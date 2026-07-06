# -*- coding: utf-8 -*-
"""Esquema vertical de columnas para elegir referencias de troceo (planos A/IA)."""

import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")

from System import AppDomain
from System.IO import File
from System import TimeSpan
from System.Windows import (
    CornerRadius,
    FontWeights,
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    TextAlignment,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    Visibility,
)
from System.Windows.Controls import (
    Border,
    Canvas,
    CheckBox,
    ColumnDefinition,
    ComboBox,
    ComboBoxItem,
    Grid,
    Orientation,
    Panel,
    RowDefinition,
    StackPanel,
    TextBlock,
)
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes, DoubleCollection, Geometry, SolidColorBrush
from System.Windows.Shapes import Ellipse, Line, Path, Rectangle
from System.Windows.Threading import DispatcherPriority, DispatcherTimer

from Autodesk.Revit.DB import FilteredElementCollector, UnitTypeId, UnitUtils
from Autodesk.Revit.DB.Structure import RebarBarType

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""

try:
    clr.AddReference("RevitAPIUI")
    from Autodesk.Revit.UI import TaskDialog
except Exception:
    TaskDialog = None


_SINGLETON_KEY = u"Arainco.column_reinforcement.TroceoSchemeSingleton"


def apply_troceo_scheme_window_maximized(win):
    u"""Despliega el formulario del esquema vertical maximizado (ventana propia o asistente)."""
    if win is None:
        return
    try:
        from System.Windows import SizeToContent, WindowState

        win.SizeToContent = SizeToContent.Manual
        win.WindowState = WindowState.Maximized
    except Exception:
        pass


def wire_wpf_numeric_stepper(tb, btn_inc, btn_dec, minimum=1, maximum=999):
    u"""Enlaza botones +/− a un ``TextBox`` con entero acotado (sigue permitiendo edición manual)."""
    if tb is None:
        return
    try:
        from System.Windows import RoutedEventHandler
    except Exception:
        return

    def read_v():
        try:
            t = tb.Text
            if t is None:
                return int(minimum)
            v = int(float(unicode(t).strip().replace(u",", u".")))
        except Exception:
            v = int(minimum)
        v = int(max(int(minimum), min(int(maximum), v)))
        return v

    def write_v(v):
        v = int(max(int(minimum), min(int(maximum), int(v))))
        tb.Text = u"{0}".format(v)

    def on_inc(sender, args):
        write_v(read_v() + 1)

    def on_dec(sender, args):
        write_v(read_v() - 1)

    if btn_inc is not None:
        btn_inc.Click += RoutedEventHandler(on_inc)
    if btn_dec is not None:
        btn_dec.Click += RoutedEventHandler(on_dec)


def _level_numeric_labels_left_px(shaft_x, shaft_w, bubble_cx, bubble_r, ls, level_strings):
    u"""X mínima (más a la izquierda) de las etiquetas de cota Z; misma regla que en ``_populate_blocks``."""
    col_right = float(shaft_x) + float(shaft_w)
    line_zone_left = col_right + max(1.5 * ls, 2.5)
    bubble_left = float(bubble_cx) - float(bubble_r)
    num_h_gap = max(2.5 * ls, 3.5)
    leftmost = None
    for s in level_strings or []:
        tb = TextBlock()
        tb.Text = s
        tb.FontSize = 10.5 * ls
        tb.FontWeight = FontWeights.SemiBold
        try:
            from System import Double
            from System.Windows import Size

            tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            lw = float(tb.DesiredSize.Width)
        except Exception:
            lw = 40.0
        lab_left = max(line_zone_left, bubble_left - num_h_gap - lw)
        if leftmost is None or lab_left < leftmost:
            leftmost = lab_left
    if leftmost is None:
        return float(line_zone_left)
    return float(leftmost)


def _hex_to_color(hex_str):
    """#RRGGBB -> System.Windows.Media.Color."""
    from System.Windows.Media import Color

    h = (hex_str or "#000000").lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return Color.FromRgb(r, g, b)


_BR_UNSEL_BG = SolidColorBrush(_hex_to_color("#0f172a"))
_BR_UNSEL_BD = SolidColorBrush(_hex_to_color("#475569"))
_BR_TEXT = SolidColorBrush(_hex_to_color("#E8F4F8"))
_BR_MUTED = SolidColorBrush(_hex_to_color("#94a3b8"))
# Relleno del fuste por tramo de troceo (1 = base); tonos oscuros coherentes con entradas/cotas BIMTools.
_TROCEO_TRAMO_SHAFT_BG_BRUSHES = tuple(
    SolidColorBrush(_hex_to_color(h))
    for h in (
        "#152535",
        "#1a3548",
        "#1d3f4a",
        "#243044",
        "#1a4038",
        "#2a3050",
        "#1f3d45",
        "#283548",
    )
)
_BR_FOUND = SolidColorBrush(_hex_to_color("#0c1824"))
# Cotas / cabeza de nivel en el esquema (alineado al tema oscuro BIMTools: menos blanco puro / azul chillón)
_BR_LEVEL_TEXT = SolidColorBrush(_hex_to_color("#A8C5D9"))
# Alerta: barra modelada supera largo comercial t\u00edpico (previsualizaci\u00f3n troceo).
_BAR_LENGTH_COMMERCIAL_MAX_MM = 12000
_BR_BAR_LENGTH_OVER_COMMERCIAL = SolidColorBrush(_hex_to_color("#f87171"))
_BR_LEVEL_LINE = SolidColorBrush(_hex_to_color("#4A6B82"))
_BR_LEVEL_DISK = SolidColorBrush(_hex_to_color("#CFDEE8"))
_BR_LEVEL_BUBBLE = SolidColorBrush(_hex_to_color("#1F7AAD"))
_BR_ENTRY_BG = SolidColorBrush(_hex_to_color("#050E18"))
_BR_ENTRY_FG = SolidColorBrush(_hex_to_color("#E8F4F8"))
_BR_ENTRY_BD = SolidColorBrush(_hex_to_color("#1A3A4D"))
# Línea de referencia troceo/empalme en el esquema (hoy: borde inferior del fuste).
_TROCEO_DATUM_REF_TAG = u"TroceoDatumRef"
_BR_TROCEO_DATUM_REF = SolidColorBrush(_hex_to_color("#dc2626"))
_TROCEO_SHAFT_TRAMO_LABEL_TAG = u"TroceoShaftTramoLabel"
_TROCEO_SHAFT_TRAMO_FILL_TAG = u"TroceoShaftTramoFill"
_LONGITUDINAL_DIAM_COMBO_TAG = u"TroceoLongitudinalDiamCombo"
_TROCEO_BAR_TYPES_MSG_TAG = u"TroceoBarTypesMsg"
# Etiquetas «Tramo n» junto al combo Ø longitudinal (se quitan con los combos).
_TROCEO_LONG_UI_DECOR_TAG = u"TroceoLongUiDecor"
# Separador vertical fuste/referencias vs columna Ø long. (solo repintado completo del lienzo).
_TROCEO_SCHEME_TROCEO_ARM_SEP_TAG = u"TroceoRefLongSep"


def _apply_troceo_scheme_header_text_style(tb, ls):
    u"""Mismo aspecto en cabecera fija (ø Estribo, Esp., Esquema, Nivel) y panel embebido."""
    if tb is None:
        return
    try:
        tb.Foreground = _BR_MUTED
        tb.FontWeight = FontWeights.SemiBold
        tb.FontSize = max(8.5, 9.0 * float(ls))
    except Exception:
        pass
    try:
        from System.Windows import FontFamily

        tb.FontFamily = FontFamily(u"Segoe UI")
    except Exception:
        pass
    try:
        from System.Windows.Media import TextFormattingMode, TextOptions

        TextOptions.SetTextFormattingMode(tb, TextFormattingMode.Display)
    except Exception:
        pass


def _rebar_type_id_int(bt):
    if bt is None:
        return None
    try:
        return int(bt.Id.Value)
    except Exception:
        try:
            return int(bt.Id.IntegerValue)
        except Exception:
            return None


def _rebar_type_nominal_mm(bt):
    if bt is None:
        return None
    try:
        return UnitUtils.ConvertFromInternalUnits(
            float(bt.BarNominalDiameter),
            UnitTypeId.Millimeters,
        )
    except Exception:
        return None


def _collect_project_rebar_bar_types(doc):
    """Lista ``(id_int, etiqueta_visible, d_mm_ord, tooltip_nombre)`` ordenada por Ø y nombre."""
    out = []
    if doc is None:
        return out
    try:
        col = FilteredElementCollector(doc).OfClass(RebarBarType)
    except Exception:
        col = []
    for bt in col:
        bid = _rebar_type_id_int(bt)
        if bid is None:
            continue
        nm = u""
        try:
            nm = (bt.Name or u"").ToString()
        except Exception:
            try:
                nm = str(bt.Name) if bt.Name is not None else u""
            except Exception:
                nm = u""
        if not nm:
            nm = u"(sin nombre)"
        dmm = _rebar_type_nominal_mm(bt)
        if dmm is None:
            label = u"Ø ?"
            dsort = 0.0
        else:
            label = u"Ø {0:.0f} mm".format(float(dmm))
            dsort = float(dmm)
        out.append((int(bid), label, dsort, nm))
    out.sort(key=lambda t: (t[2], t[3], t[0]))
    return out


def _default_combo_index_for_mm(choices, target_mm):
    """Índice del RebarBarType cuyo Ø nominal está más cerca de ``target_mm``."""
    if not choices:
        return 0
    best_i = 0
    best_d = None
    tt = float(target_mm)
    for i, row in enumerate(choices):
        dm = float(row[2])
        delta = abs(dm - tt)
        if best_d is None or delta < best_d:
            best_d = delta
            best_i = i
    return int(best_i)


def _longitudinal_diam_tags_resized(base, n_segs):
    u"""Adapta la lista de Id de RebarBarType (un tramo = un \u00edndice) al nuevo n\u00famero de bandas."""
    n = max(0, int(n_segs))
    if n == 0:
        return []
    if not base:
        return [None] * n
    out = []
    for t in base:
        try:
            out.append(int(t) if t is not None else None)
        except Exception:
            out.append(None)
    if len(out) >= n:
        return out[:n]
    last_good = None
    for t in reversed(out):
        if t is not None:
            last_good = t
            break
    fill = last_good
    if fill is None and out:
        fill = out[-1]
    while len(out) < n:
        out.append(fill)
    return out


def _unpack_troceo_row(row):
    """Acepta 3- a 6-tuplas desde el layout legado o actual."""
    if not row:
        return None, 0.0, -1, None, None, None
    try:
        n = len(row)
        if n >= 6:
            e, z, i = row[0], float(row[1]), int(row[2])
            return e, z, i, row[3], row[4], row[5]
        if n >= 5:
            e, z, i = row[0], float(row[1]), int(row[2])
            return e, z, i, row[3], row[4], None
        if n == 4:
            e, z, i = row[0], float(row[1]), int(row[2])
            return e, z, i, row[3], None, None
        e, z, i = row[0], float(row[1]), int(row[2])
        return e, z, i, None, None, None
    except Exception:
        return None, 0.0, -1, None, None, None


def _column_plan_dims_short_long_mm(elem):
    """Devuelve ``(corto_mm, largo_mm)`` de sección en planta o ``(None, None)``."""
    if elem is None:
        return None, None
    try:
        from column_reinforcement_layout_rps import get_column_dimensions

        w_ft, d_ft, _, _, _, _ = get_column_dimensions(elem)
        w_mm = UnitUtils.ConvertFromInternalUnits(
            float(w_ft),
            UnitTypeId.Millimeters,
        )
        d_mm = UnitUtils.ConvertFromInternalUnits(
            float(d_ft),
            UnitTypeId.Millimeters,
        )
        ss = min(float(w_mm), float(d_mm))
        sl = max(float(w_mm), float(d_mm))
        return int(round(ss)), int(round(sl))
    except Exception:
        return None, None


def _add_left_segment_dims_vertical(
    canvas,
    shaft_left_x,
    y_top,
    hpx,
    ss,
    sl,
    z_index,
    layout_scale=1.0,
    section_label_left_x=None,
    section_zone_w=None,
):
    u"""Sección ``400×600`` rotada −90°; centrada en su columna si se indica ``section_label_left_x``."""
    try:
        from System import Double
        from System.Windows import Size
        from System.Windows.Media import RotateTransform
    except Exception:
        return
    if ss is None or sl is None:
        return
    ls = float(layout_scale)
    cap = u"{0}x{1}".format(int(ss), int(sl))
    tb = TextBlock()
    tb.Text = cap
    base_fs = 8.25 if hpx >= 40.0 else 7.5
    tb.FontSize = base_fs * ls
    tb.Foreground = _BR_MUTED
    tb.FontWeight = FontWeights.SemiBold
    try:
        tb.LayoutTransform = RotateTransform(-90.0)
    except Exception:
        return
    try:
        tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
        box_w = float(tb.DesiredSize.Width)
        box_h = float(tb.DesiredSize.Height)
    except Exception:
        box_w = 80.0
        box_h = 14.0
    pad = _TROCEO_SECTION_LABEL_PAD_PX * ls
    try:
        if section_label_left_x is not None and section_zone_w is not None:
            zone_l = float(section_label_left_x)
            zone_w = max(float(section_zone_w), box_w + 2.0 * pad)
            left_x = zone_l + max(pad, (zone_w - box_w) / 2.0)
        else:
            left_x = max(1.0, float(shaft_left_x) - box_w - pad)
        Canvas.SetLeft(tb, left_x)
        Canvas.SetTop(tb, y_top + max((hpx - box_h) / 2.0, 0.0))
        Panel.SetZIndex(tb, int(z_index))
    except Exception:
        pass
    canvas.Children.Add(tb)


def _finalize_segment_heights(models):
    """Rellena ``height_mm`` (mm) y etiqueta de nivel por defecto."""
    n = len(models)
    if n == 0:
        return
    zs = [float(p[u"z_mm"]) for p in models]
    for i, p in enumerate(models):
        h = p.get(u"height_mm")
        try:
            h = float(h) if h is not None else None
        except Exception:
            h = None
        if h is None or h < 100.0:
            if i < n - 1:
                dz = zs[i + 1] - zs[i]
                h = dz if dz > 50.0 else 2800.0
            elif n >= 2:
                dz = zs[i] - zs[i - 1]
                h = dz if dz > 50.0 else 2800.0
            else:
                h = 3000.0
        p[u"height_mm"] = float(max(h, 200.0))
        if not p.get(u"level_label"):
            zs_i = zs[i]
            p[u"level_label"] = u"{0:.3f}".format(float(zs_i) / 1000.0)


def _shaft_brushes_for_tramo(tramo_no):
    u"""``tramo_no`` 1-based como en la escalerilla de barras (tramo 1 = base)."""
    bl = _TROCEO_TRAMO_SHAFT_BG_BRUSHES
    idx = (max(1, int(tramo_no)) - 1) % len(bl)
    return bl[idx], _BR_UNSEL_BD


def _tramo_number_at_canvas_y(y_px, bands):
    u"""Tramo 1-based que contiene la cota Y en pantalla."""
    try:
        y = float(y_px)
    except Exception:
        return 1
    for y_top_cell, h_cell, tramo_no in bands or []:
        try:
            y0 = float(y_top_cell)
            y1 = y0 + float(h_cell)
        except Exception:
            continue
        if y0 - 1e-6 <= y <= y1 + 1e-6:
            return int(tramo_no)
    return 1


def _remove_troceo_shaft_tramo_fills(canvas):
    if canvas is None:
        return
    to_remove = []
    try:
        for ch in canvas.Children:
            try:
                if ch.Tag == _TROCEO_SHAFT_TRAMO_FILL_TAG:
                    to_remove.append(ch)
            except Exception:
                pass
        for ch in to_remove:
            canvas.Children.Remove(ch)
    except Exception:
        pass


def _paint_troceo_shaft_tramo_fills(canvas, shaft_x, shaft_w, slot_to_y_span, bands, cut_ys):
    u"""
    Relleno del fuste por subtramos entre l\u00edneas de corte (incl. mitad de altura dentro de un slot).
    """
    if canvas is None or not slot_to_y_span:
        return
    try:
        sx = float(shaft_x)
        sw = max(4.0, float(shaft_w))
    except Exception:
        return
    cuts_all = []
    for c in cut_ys or []:
        try:
            cuts_all.append(float(c))
        except Exception:
            pass
    _remove_troceo_shaft_tramo_fills(canvas)
    for _sl, span in (slot_to_y_span or {}).items():
        if not span or len(span) < 2:
            continue
        try:
            y_top = float(span[0])
            y_bot = float(span[1])
        except Exception:
            continue
        if y_bot <= y_top + 0.5:
            continue
        interior = [y_top]
        for c in cuts_all:
            if y_top + 1e-3 < c < y_bot - 1e-3:
                interior.append(c)
        interior.append(y_bot)
        bnds = sorted(set(interior))
        for i in range(len(bnds) - 1):
            ya = float(bnds[i])
            yb = float(bnds[i + 1])
            dz = yb - ya
            if dz < 0.5:
                continue
            ym = ya + 0.5 * dz
            tno = _tramo_number_at_canvas_y(ym, bands)
            bg, _ = _shaft_brushes_for_tramo(tno)
            rect = Rectangle()
            rect.Width = sw
            rect.Height = dz
            rect.Fill = bg
            try:
                rect.Stroke = None
            except Exception:
                pass
            try:
                rect.Tag = _TROCEO_SHAFT_TRAMO_FILL_TAG
            except Exception:
                pass
            Canvas.SetLeft(rect, sx)
            Canvas.SetTop(rect, ya)
            try:
                Panel.SetZIndex(rect, 11)
            except Exception:
                pass
            try:
                canvas.Children.Add(rect)
            except Exception:
                pass


def _slot_to_tramo_number_from_bands(slot, bands, slot_to_y_span):
    u"""Tramo (1 = base) del segmento ``slot`` seg\u00fan bandas de ``_bar_tramo_y_bands_from_cuts``."""
    span = slot_to_y_span.get(int(slot)) if slot_to_y_span else None
    if not span or not bands:
        return 1
    try:
        yc = (float(span[0]) + float(span[1])) / 2.0
    except Exception:
        return 1
    for y_top_cell, h_cell, tramo_no in bands:
        try:
            y0 = float(y_top_cell)
            y1 = y0 + float(h_cell)
        except Exception:
            continue
        if y0 - 1e-6 <= yc <= y1 + 1e-6:
            return int(tramo_no)
    return 1


def _shaft_segment_border_and_radius(slot, n_seg, ls, selected):
    u"""Borde sin duplicar trazos horizontales entre tramos (apilado continuo); redondeo solo en el contorno."""
    ls = float(ls)
    t = max(1.0, (2.0 if selected else 1.2) * ls)
    r = max(1.0, 1.0 * ls)
    n_seg = int(n_seg)
    slot = int(slot)
    if n_seg <= 1:
        return Thickness(t), CornerRadius(r)
    is_top = slot == n_seg - 1
    is_bot = slot == 0
    th = Thickness(t, t if is_top else 0.0, t, t if is_bot else 0.0)
    if is_top:
        cr = CornerRadius(r, r, 0.0, 0.0)
    elif is_bot:
        cr = CornerRadius(0.0, 0.0, r, r)
    else:
        cr = CornerRadius(0.0, 0.0, 0.0, 0.0)
    return th, cr


def _draw_foundation(canvas, fx, fy, fw, fh):
    rect = Rectangle()
    rect.Width = fw
    rect.Height = fh
    rect.Fill = _BR_FOUND
    rect.Stroke = _BR_MUTED
    rect.StrokeThickness = 1.0
    Canvas.SetLeft(rect, fx)
    Canvas.SetTop(rect, fy)
    canvas.Children.Add(rect)


def _add_troceo_datum_reference_line(canvas, y_ref, x1, x2, layout_scale, z_index=13):
    u"""Marca la cota de referencia usada para troceo y empalme (hoy: borde inferior del fuste)."""
    if canvas is None:
        return
    try:
        ls = float(layout_scale)
    except Exception:
        ls = 1.0
    ln = Line()
    ln.X1 = float(x1)
    ln.X2 = float(max(float(x2), float(x1) + 1.0))
    y = float(y_ref)
    ln.Y1 = y
    ln.Y2 = y
    ln.Stroke = _BR_TROCEO_DATUM_REF
    ln.StrokeThickness = max(1.35, 1.6 * ls)
    try:
        ln.Tag = _TROCEO_DATUM_REF_TAG
    except Exception:
        pass
    try:
        Panel.SetZIndex(ln, int(z_index))
    except Exception:
        pass
    canvas.Children.Add(ln)


def _remove_troceo_datum_reference_lines(canvas):
    u"""Quita l\u00edneas de referencia troceo previas (al cambiar selecci\u00f3n sin repintar todo)."""
    if canvas is None:
        return
    try:
        to_remove = []
        for ch in canvas.Children:
            try:
                if ch.Tag == _TROCEO_DATUM_REF_TAG:
                    to_remove.append(ch)
            except Exception:
                pass
        for ch in to_remove:
            canvas.Children.Remove(ch)
    except Exception:
        pass


def _remove_troceo_shaft_tramo_labels(canvas):
    u"""Quita n\u00fameros de tramo dibujados sobre el fuste."""
    if canvas is None:
        return
    try:
        to_remove = []
        for ch in canvas.Children:
            try:
                if ch.Tag == _TROCEO_SHAFT_TRAMO_LABEL_TAG:
                    to_remove.append(ch)
            except Exception:
                pass
        for ch in to_remove:
            canvas.Children.Remove(ch)
    except Exception:
        pass


def _draw_troceo_shaft_tramo_labels(canvas, shaft_x, shaft_w, bands, layout_scale):
    u"""N\u00famero 1\u2026N centrado en cada banda de troceo dentro del ancho del fuste."""
    if canvas is None or not bands:
        return
    try:
        from System import Double
        from System.Windows import Size
        from System.Windows.Controls import ToolTip
    except Exception:
        ToolTip = None
        Double = None
        Size = None
    ls = float(layout_scale)
    sx = float(shaft_x)
    sw = max(4.0, float(shaft_w))
    zx = 16
    tag = _TROCEO_SHAFT_TRAMO_LABEL_TAG
    for y_top_cell, h_cell, tramo_no in bands:
        n = int(tramo_no)
        tb = TextBlock()
        tb.Text = u"{0}".format(n)
        tb.Foreground = _BR_LEVEL_TEXT
        tb.FontSize = max(10.5, 12.5 * ls)
        tb.FontWeight = FontWeights.SemiBold
        tb.TextAlignment = TextAlignment.Center
        try:
            tb.Width = sw
        except Exception:
            pass
        tb.Tag = tag
        if ToolTip is not None:
            try:
                tip = ToolTip()
                tip.Content = u"Tramo {0} de troceo (1 = base del pilar).".format(n)
                tb.ToolTip = tip
            except Exception:
                pass
        try:
            if Size is not None and Double is not None:
                tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
                th = float(tb.DesiredSize.Height)
            else:
                th = max(14.0, 13.0 * ls)
        except Exception:
            th = max(14.0, 13.0 * ls)
        yt = float(y_top_cell) + max((float(h_cell) - th) / 2.0, 0.0)
        Canvas.SetLeft(tb, sx)
        Canvas.SetTop(tb, yt)
        try:
            Panel.SetZIndex(tb, zx)
        except Exception:
            pass
        try:
            canvas.Children.Add(tb)
        except Exception:
            pass


def _remove_longitudinal_diam_scheme_widgets(canvas):
    u"""Quita combos de \u00d8 longitudinal y mensajes de error asociados del canvas del esquema."""
    if canvas is None:
        return
    try:
        to_remove = []
        for ch in canvas.Children:
            try:
                t = ch.Tag
                if t in (
                    _LONGITUDINAL_DIAM_COMBO_TAG,
                    _TROCEO_BAR_TYPES_MSG_TAG,
                    _TROCEO_LONG_UI_DECOR_TAG,
                ):
                    to_remove.append(ch)
            except Exception:
                pass
        for ch in to_remove:
            canvas.Children.Remove(ch)
    except Exception:
        pass


def _add_troceo_ref_long_vertical_separator(canvas, x, y0, y1, layout_scale):
    u"""L\u00ednea vertical sutil entre referencias de troceo y columna de \u00d8 longitudinal."""
    if canvas is None:
        return
    try:
        from System.Windows.Shapes import Line

        ls = float(layout_scale)
        ln = Line()
        xx = float(x)
        ln.X1 = xx
        ln.X2 = xx
        ln.Y1 = float(y0)
        ln.Y2 = float(y1)
        try:
            ln.Stroke = _BR_ENTRY_BD
            ln.Opacity = 0.55
        except Exception:
            pass
        ln.StrokeThickness = max(1.0, 1.0 * ls)
        ln.Tag = _TROCEO_SCHEME_TROCEO_ARM_SEP_TAG
        try:
            Panel.SetZIndex(ln, 4)
        except Exception:
            pass
        canvas.Children.Add(ln)
    except Exception:
        pass


def _apply_troceo_band_tramo_label_style(tb, ls):
    u"""Etiqueta de tramo de troceo (n\u00famero de banda) junto al combo longitudinal."""
    if tb is None:
        return
    try:
        tb.Foreground = _BR_MUTED
        tb.FontWeight = FontWeights.Normal
        tb.FontSize = max(7.5, 7.85 * float(ls))
    except Exception:
        pass
    try:
        from System.Windows import FontFamily

        tb.FontFamily = FontFamily(u"Segoe UI")
    except Exception:
        pass


# Por debajo de cotas rotadas (3), combos (8–9) y fuste (12): atraviesa el lienzo sin tapar etiquetas.
_Z_TROCEO_DATUM_LINE = 2


_BAR_TRAMO_LADDER_TAG = u"BarTramoLadderShape"
_BAR_LENGTH_EST_TAG = u"BarLengthEst"

# Escala del esquema: con pocos tramos se usa la misma altura en pantalla que si hubiera al menos N (mejor lectura).
_MIN_TROCEO_SEGMENTS_FOR_SCALE = 4
# Tope de la pila de tramos (px base, antes de × layout_scale); subir para pantallas altas / fase troceo a pantalla completa.
_MAX_TROCEO_STACK_PX = 2000.0
_TROCEO_SEG_UNIFORM_BASE_PX = 120.0
_TROCEO_SEG_UNIFORM_MIN_PX = 62.0
# Factor único para agrandar fuste, escalerilla, cimentación y cotas en pantalla.
_TROCEO_LAYOUT_SCALE = 1.62
# Ancho nominal (× layout_scale) del combo en fila; alto se calcula aparte, más bajo.
_TROCEO_STIRRUP_COMBO_BAR_TYPE_W_PX = 58.0
# Solo si no hay combo de tipo (caso raro).
_TROCEO_STIRRUP_COMBO_SPACING_W_PX = 46.0
_TROCEO_STIRRUP_COMBO_PAIR_GAP_PX = 6.0
# Más separación del fuste para estribos + etiqueta de sección + política (legibilidad).
_TROCEO_SHAFT_X_BASE_PX = 96.0
_TROCEO_SHAFT_LABEL_MARGIN_PX = 10.0
_TROCEO_SECTION_LABEL_PAD_PX = 10.0
_TROCEO_POLICY_SECTION_GAP_PX = 14.0
_TROCEO_STIRRUP_INTER_COL_GAP_PX = 8.0
_TROCEO_EMPALME_POLICY_COL_W_PX = 100.0
# Mismos valores que column_reinforcement_layout_rps (troceo_empalme_policy_by_column_id).
TROCEO_EMPALME_POLICY_BASE = u"base"
TROCEO_EMPALME_POLICY_MID_AXIS = u"mid_axis_split"
_TROCEO_EMPALME_POLICY_UI_CHOICES = (
    (u"Base", TROCEO_EMPALME_POLICY_BASE),
    (u"Mitad altura", TROCEO_EMPALME_POLICY_MID_AXIS),
)


def _troceo_empalme_policy_combo_width_px(layout_scale):
    return max(float(_TROCEO_EMPALME_POLICY_COL_W_PX) * float(layout_scale), 88.0)


def _stirrup_combo_row_width_px(layout_scale, has_bar_type_combo=True):
    ls = float(layout_scale)
    if has_bar_type_combo:
        return max(float(_TROCEO_STIRRUP_COMBO_BAR_TYPE_W_PX) * ls, 24.0)
    return max(float(_TROCEO_STIRRUP_COMBO_SPACING_W_PX) * ls, 24.0)


def _stirrup_combo_compact_height_px(layout_scale):
    u"""Alto del combo ajustado al texto (sin igualar al ancho). Incluye margen para padding vertical."""
    ls = float(layout_scale)
    return max(22.0, 13.0 * ls + 4.0)


def _stirrup_combo_vertical_padding_px(layout_scale):
    u"""Padding vertical interior de los combos estribo (esquema y panel Tramos)."""
    ls = float(layout_scale)
    return max(3.0, 2.5 * ls)

# Opciones de espaciamiento (mm) en combos incrustados del esquema (asistente).
_STIRRUP_SPACING_MM_CHOICES = (
    u"100",
    u"125",
    u"150",
    u"175",
    u"200",
    u"225",
    u"250",
    u"300",
)

STIRRUP_POLICY_CONTINUOUS = u"continuous"
STIRRUP_POLICY_THIRDS_L3 = u"thirds_l3"
_STIRRUP_POLICY_UI_CHOICES = (
    (u"Completo", STIRRUP_POLICY_CONTINUOUS),
    (u"L/3", STIRRUP_POLICY_THIRDS_L3),
)
_STIRRUP_LOT_LABELS = (u"T1", u"T2", u"T3")
# Espaciamiento por defecto en extremos al activar política L/3 (T1 inferior, T3 superior).
_STIRRUP_L3_T1_T3_DEFAULT_SPACING_MM = 100.0


def _stirrup_slot_lot_store_key(slot, lot_index=0):
    u"""Clave compuesta tramo esquema × lote vertical (0=T1 inferior … 2=T3)."""
    return int(slot) * 3 + int(lot_index)


def _stirrup_policy_for_slot(policy_store, slot):
    if policy_store is None:
        return STIRRUP_POLICY_CONTINUOUS
    try:
        if policy_store.get(int(slot)) == STIRRUP_POLICY_THIRDS_L3:
            return STIRRUP_POLICY_THIRDS_L3
    except Exception:
        pass
    return STIRRUP_POLICY_CONTINUOUS


def _stirrup_spacing_from_store(store, store_key, slot, elem, default_cb):
    if store is None:
        return None
    try:
        if int(store_key) in store:
            return float(store[int(store_key)])
    except Exception:
        pass
    try:
        if int(store_key) == _stirrup_slot_lot_store_key(slot, 0) and int(slot) in store:
            return float(store[int(slot)])
    except Exception:
        pass
    if default_cb is not None and elem is not None:
        try:
            return float(default_cb(elem))
        except Exception:
            pass
    return None


def _stirrup_bar_type_from_store(store, store_key, slot, elem, default_cb, choices):
    if store is None:
        return None
    bt = None
    try:
        if int(store_key) in store:
            bt = store[int(store_key)]
    except Exception:
        bt = None
    if bt is None:
        try:
            if int(slot) in store:
                bt = store[int(slot)]
        except Exception:
            bt = None
    if bt is None and default_cb is not None and elem is not None:
        try:
            bt = default_cb(elem)
        except Exception:
            bt = None
    return bt


def _stirrup_policy_combo_width_px(layout_scale):
    u"""Ancho fijo del combo «Completo» / «L/3» (evita invadir la columna de sección)."""
    return max(58.0, 54.0 * float(layout_scale))


def _stirrup_lote_label_col_width_px(layout_scale):
    return max(26.0, 22.0 * float(layout_scale))


def _measure_section_label_bounds(dims_cache, layout_scale):
    u"""Ancho/alto del bloque «400×600» rotado −90° (mismo criterio que en el lienzo)."""
    ls = float(layout_scale)
    max_w = 0.0
    max_h = 0.0
    try:
        from System import Double
        from System.Windows import Size
        from System.Windows.Controls import TextBlock
        from System.Windows.Media import RotateTransform
        from System.Windows import FontWeights
    except Exception:
        return max(48.0 * ls, 36.0), max(14.0 * ls, 12.0)
    for _eid, pair in (dims_cache or {}).items():
        ss, sl = pair
        if ss is None or sl is None:
            continue
        cap = u"{0}x{1}".format(int(ss), int(sl))
        for base_fs in (7.5 * ls, 8.25 * ls):
            tb = TextBlock()
            tb.Text = cap
            tb.FontSize = float(base_fs)
            tb.FontWeight = FontWeights.SemiBold
            tb.LayoutTransform = RotateTransform(-90.0)
            try:
                tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
                max_w = max(max_w, float(tb.DesiredSize.Width))
                max_h = max(max_h, float(tb.DesiredSize.Height))
            except Exception:
                pass
    if max_w < 1.0:
        max_w = 48.0 * ls
    if max_h < 1.0:
        max_h = 14.0 * ls
    return max_w, max_h


def _troceo_stirrup_strip_x_layout(shaft_x, dims_cache, layout_scale, combo_w_px):
    u"""
    Columnas fijas (izq. → der.): Lote | Ø+Esp. | Política | Sección | fuste.
    Devuelve posiciones X en el mismo sistema que ``shaft_x``.
    """
    ls = float(layout_scale)
    max_rot_w, max_rot_h = _measure_section_label_bounds(dims_cache, ls)
    pad = _TROCEO_SECTION_LABEL_PAD_PX * ls
    shaft_margin = _TROCEO_SHAFT_LABEL_MARGIN_PX * ls
    pol_sec_gap = _TROCEO_POLICY_SECTION_GAP_PX * ls
    inter_col = _TROCEO_STIRRUP_INTER_COL_GAP_PX * ls
    lote_gap = max(5.0, 4.5 * ls)

    section_w = max(
        float(max_rot_w) + 2.0 * pad,
        float(max_rot_h) + 2.0 * pad,
        52.0 * ls,
    )
    policy_w = _stirrup_policy_combo_width_px(ls)
    lote_w = _stirrup_lote_label_col_width_px(ls)

    section_right = float(shaft_x) - shaft_margin
    section_left = section_right - section_w
    policy_right = section_left - pol_sec_gap
    policy_x = policy_right - policy_w
    combo_right = policy_x - inter_col
    combo_left = combo_right - float(combo_w_px)
    lote_col_x = combo_left - lote_gap - lote_w

    return {
        u"section_left": section_left,
        u"section_right": section_right,
        u"section_w": section_w,
        u"policy_x": policy_x,
        u"policy_w": policy_w,
        u"combo_left_x": combo_left,
        u"lote_col_x": lote_col_x,
        u"strip_left": lote_col_x,
    }


def _bar_tramo_equal_bands(y_top_col, y_bot, m):
    """Reparto vertical uniforme entre ``y_top_col`` y ``y_bot`` (respaldo). Orden: tramo 1 abajo."""
    m = max(1, int(m))
    y_top_col = float(y_top_col)
    y_bot = float(y_bot)
    total = y_bot - y_top_col
    if total < 1.0:
        return [(y_top_col, max(4.0, total), 1)]
    h = total / float(m)
    bands = []
    y_cursor = y_bot
    for j in range(m):
        y_top_cell = y_cursor - h
        bands.append((y_top_cell, h, j + 1))
        y_cursor = y_top_cell
    return bands


def _bar_tramo_y_bands_from_cuts(
    sel_slots, slot_to_y_bottom, top_pad, stack_h, slot_to_cut_y=None
):
    u"""Bandas del fuste entre cotas de troceo/empalme (Y base o mitad de tramo por referencia)."""
    y_bot = float(top_pad) + float(stack_h)
    y_top_col = float(top_pad)
    m = max(1, len(sel_slots) + 1)
    if m == 1:
        return _bar_tramo_equal_bands(y_top_col, y_bot, 1)
    eps = 0.5
    cuts = []
    cut_src = slot_to_cut_y if slot_to_cut_y else {}
    for s in sel_slots or []:
        try:
            ii = int(s)
            yy = None
            if ii in cut_src:
                yy = float(cut_src[ii])
            elif ii in slot_to_y_bottom:
                yy = float(slot_to_y_bottom[ii])
            if yy is not None and (y_top_col + eps) <= yy <= (y_bot - eps):
                cuts.append(yy)
        except Exception:
            pass
    cuts = sorted(set(cuts))
    bnds = [y_top_col] + cuts + [y_bot]
    if len(bnds) != m + 1:
        return _bar_tramo_equal_bands(y_top_col, y_bot, m)
    bands = []
    for j in range(m):
        y0 = float(bnds[m - 1 - j])
        y1 = float(bnds[m - j])
        if y1 < y0 + 1e-6:
            return _bar_tramo_equal_bands(y_top_col, y_bot, m)
        bands.append((y0, y1 - y0, j + 1))
    return bands


def _troceo_long_combo_top_px_for_band(y_top_cell, h_cell, slot_span, h_cb):
    u"""Y superior del combo \u00d8 longitudinal alineado con una fila f\u00edsica de la banda.

    Si hay varias filas (sub-tramos), se usa la fila **inferior** de la banda en pantalla
    (mayor Y, cercana al corte con el tramo de troceo siguiente hacia la base): as\u00ed el
    combo comparte l\u00ednea con una casilla «Define Empalme» y no queda flotando entre dos.
    Una sola fila: mismo resultado que centrar en esa fila.
    """
    y0b = float(y_top_cell)
    h = float(h_cell)
    y1b = y0b + h
    hh = float(h_cb)
    rows = []
    if slot_span:
        try:
            for _sl, span in slot_span.items():
                if not span or len(span) < 2:
                    continue
                ys = float(span[0])
                ye = float(span[1])
                if ye <= y0b + 1e-6 or ys >= y1b - 1e-6:
                    continue
                y_mid = ys + (ye - ys) / 2.0
                rows.append((ye, y_mid))
        except Exception:
            rows = []
    if rows:
        rows.sort(key=lambda t: t[0], reverse=True)
        y_mid = float(rows[0][1])
        top = y_mid - hh / 2.0
        return max(y0b, min(top, y1b - hh))
    return y0b + max((h - hh) / 2.0, 0.0)


def _troceo_anchor_slot_index_for_band(y_top_cell, h_cell, slot_span):
    u"""Slot f\u00edsico ancla del combo \u00d8 long.: fila con mayor Y dentro de la banda (misma l\u00f3gica que el combo)."""
    y0b = float(y_top_cell)
    y1b = y0b + float(h_cell)
    best_sl = None
    best_ye = None
    if not slot_span:
        return None
    try:
        for sl, span in slot_span.items():
            if not span or len(span) < 2:
                continue
            ys = float(span[0])
            ye = float(span[1])
            if ye <= y0b + 1e-6 or ys >= y1b - 1e-6:
                continue
            if best_ye is None or ye > float(best_ye):
                best_ye = ye
                best_sl = int(sl)
    except Exception:
        return None
    return best_sl


def _remove_bar_tramo_ladder_shapes(canvas):
    if canvas is None:
        return
    try:
        coll = canvas.Children
        rm = []
        for i in range(int(coll.Count)):
            ch = coll[i]
            try:
                if ch.Tag == _BAR_TRAMO_LADDER_TAG:
                    rm.append(ch)
            except Exception:
                pass
        for ch in rm:
            coll.Remove(ch)
    except Exception:
        pass


def _remove_bar_length_est_shapes(canvas):
    if canvas is None:
        return
    try:
        coll = canvas.Children
        rm = []
        for i in range(int(coll.Count)):
            ch = coll[i]
            try:
                if ch.Tag == _BAR_LENGTH_EST_TAG:
                    rm.append(ch)
            except Exception:
                pass
        for ch in rm:
            coll.Remove(ch)
    except Exception:
        pass


def _bar_length_total_mm_from_preview_text(text):
    u"""Extrae el entero despu\u00e9s de ``L=`` en etiquetas tipo ``A L=12345 (..)``."""
    if text is None:
        return None
    try:
        s = unicode(text)
    except Exception:
        try:
            s = u"{0}".format(text)
        except Exception:
            return None
    lu = s.upper()
    idx = lu.find(u"L=")
    if idx < 0:
        return None
    j = idx + 2
    n = len(s)
    while j < n and s[j].isspace():
        j += 1
    if j >= n or not s[j].isdigit():
        return None
    k = j
    while k < n and s[k].isdigit():
        k += 1
    try:
        return int(s[j:k])
    except Exception:
        return None


def _apply_bar_length_preview_foreground(tb, display_text):
    u"""Rojo si el largo total > l\u00edmite comercial; si no, color de cota del esquema."""
    if tb is None:
        return
    tot = _bar_length_total_mm_from_preview_text(display_text)
    try:
        if tot is not None and tot > _BAR_LENGTH_COMMERCIAL_MAX_MM:
            tb.Foreground = _BR_BAR_LENGTH_OVER_COMMERCIAL
        else:
            tb.Foreground = _BR_LEVEL_TEXT
    except Exception:
        pass


def _position_vertical_bar_length_label(tb, y0, h_cell, bar_x0_ladder, ls_f):
    u"""Texto rotado −90° (vertical), pegado a la izquierda de la escalerilla."""
    if tb is None:
        return
    try:
        from System import Double
        from System.Windows import Size
        from System.Windows.Media import RotateTransform
    except Exception:
        return
    try:
        tb.LayoutTransform = RotateTransform(-90.0)
    except Exception:
        return
    try:
        tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
        box_w = float(tb.DesiredSize.Width)
        box_h = float(tb.DesiredSize.Height)
    except Exception:
        box_w = 18.0
        box_h = 14.0
    pad = max(2.0, 3.0 * float(ls_f))
    try:
        Canvas.SetLeft(tb, max(0.0, float(bar_x0_ladder) - box_w - pad))
        Canvas.SetTop(tb, float(y0) + max((float(h_cell) - box_h) / 2.0, 0.0))
    except Exception:
        pass


def _apply_side_by_side_vertical_bar_length_geometry(
    row_tbs,
    y0,
    h_cell,
    bar_x0_ladder,
    ls_f,
):
    u"""Varias etiquetas (A, B, \u2026): texto vertical, en fila horizontal hacia la izquierda (primera m\u00e1s cerca de la escalerilla)."""
    if not row_tbs:
        return
    try:
        from System import Double
        from System.Windows import Size
        from System.Windows.Media import RotateTransform
    except Exception:
        return
    ls_f = float(ls_f)
    y0 = float(y0)
    h_cell = float(h_cell)
    bx = float(bar_x0_ladder)
    pad = max(2.0, 3.0 * ls_f)
    gap = max(1.5, 2.0 * ls_f)
    x_right = bx - pad
    for tb in row_tbs:
        if tb is None:
            continue
        try:
            tb.LayoutTransform = RotateTransform(-90.0)
            tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            box_w = float(tb.DesiredSize.Width)
            box_h = float(tb.DesiredSize.Height)
        except Exception:
            box_w = max(16.0, 14.0 * ls_f)
            box_h = max(12.0, 10.0 * ls_f)
        try:
            Canvas.SetLeft(tb, max(0.0, x_right - box_w))
            Canvas.SetTop(tb, y0 + max((h_cell - box_h) / 2.0, 0.0))
        except Exception:
            pass
        x_right = x_right - box_w - gap


def _draw_bar_tramos_ladder_cells(
    canvas,
    bar_x0,
    top_y,
    stack_h,
    bands,
    bar_w,
    z_base=0,
    layout_scale=1.0,
):
    """``bands``: ``(y_superior, alto_px, nº tramo)``; tramo 1 = base (primera entrada)."""
    if not bands:
        return
    try:
        from System import Double
        from System.Windows import Size
    except Exception:
        return
    ls = float(layout_scale)
    stk = max(1.0, 1.0 * ls)
    zx = int(z_base)
    bx = float(bar_x0)
    bw = float(bar_w)
    ty = float(top_y)
    sh = float(stack_h)
    tag = _BAR_TRAMO_LADDER_TAG

    outer = Rectangle()
    outer.Width = bw
    outer.Height = sh
    outer.Fill = Brushes.Transparent
    outer.Stroke = _BR_TEXT
    outer.StrokeThickness = stk
    Canvas.SetLeft(outer, bx)
    Canvas.SetTop(outer, ty)
    outer.Tag = tag
    try:
        Panel.SetZIndex(outer, zx)
    except Exception:
        pass
    canvas.Children.Add(outer)

    for j in range(len(bands) - 1):
        y_line = float(bands[j][0])
        ln = Line()
        ln.X1 = bx
        ln.X2 = bx + bw
        ln.Y1 = y_line
        ln.Y2 = y_line
        ln.Stroke = _BR_TEXT
        ln.StrokeThickness = stk
        ln.Tag = tag
        try:
            Panel.SetZIndex(ln, zx + 1)
        except Exception:
            pass
        canvas.Children.Add(ln)

    for y_top_cell, h_cell, tramo_no in bands:
        tb = TextBlock()
        tb.Text = u"{0}".format(int(tramo_no))
        tb.Foreground = _BR_TEXT
        tb.FontSize = 11.0 * ls
        tb.FontWeight = FontWeights.SemiBold
        tb.TextAlignment = TextAlignment.Center
        tb.Width = bw
        tb.Tag = tag
        try:
            tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            th = float(tb.DesiredSize.Height)
        except Exception:
            th = 14.0
        yt = float(y_top_cell) + max((float(h_cell) - th) / 2.0, 0.0)
        Canvas.SetLeft(tb, bx)
        Canvas.SetTop(tb, yt)
        try:
            Panel.SetZIndex(tb, zx + 2)
        except Exception:
            pass
        canvas.Children.Add(tb)


def _tramo_band_row_heights_top_to_bottom_px(bands, n_segs, stack_h):
    u"""Pixeles por fila de selector: arriba tramo N … abajo tramo 1 (coherente con ``bands``)."""
    ns = max(1, int(n_segs))
    sh = float(stack_h)
    if bands and len(bands) == ns:
        out = []
        for k in range(ns - 1, -1, -1):
            try:
                _y0, h_cell, _tno = bands[k]
                out.append(max(28.0, float(h_cell)))
            except Exception:
                out.append(max(36.0, sh / float(ns)))
        return out
    eq = max(36.0, sh / float(ns))
    return [eq] * ns


def _slots_overlapping_band_px(band_y_top, band_h, slot_to_y_span, eps=0.5):
    u"""Slots físicos cuyo trazo vertical intersecta la banda del tramo (coords. canvas, Y hacia abajo)."""
    if not slot_to_y_span:
        return []
    y1 = float(band_y_top)
    y2 = float(band_y_top) + float(band_h)
    out = []
    for sl, span in slot_to_y_span.items():
        try:
            st, sb = float(span[0]), float(span[1])
        except Exception:
            continue
        top = min(st, sb)
        bot = max(st, sb)
        t0 = max(y1, top)
        t1 = min(y2, bot)
        if t1 - t0 > float(eps):
            out.append(int(sl))
    return sorted(set(out))


def _nominal_mm_from_bar_combo(cb, choices, default_mm):
    u"""Ø nominal (mm) del ítem seleccionado en un combo de barras."""
    try:
        sel = cb.SelectedItem
        if sel is None:
            return float(default_mm)
        tid = int(sel.Tag)
        for bid, _lbl, dmm, _nm in choices or []:
            if int(bid) == tid:
                return float(dmm)
    except Exception:
        pass
    return float(default_mm)


def _format_long_bar_length_estimate_mm(
    seg_models,
    band_slot_indices,
    diameter_mm,
    concrete_grade,
    doc=None,
    bar_enum_lab=None,
    n_total_tramos=1,
    top_collides=None,
):
    u"""Largo por l\u00ednea de fierro: tronco + empalmes + extremos, alineado al pipeline del layout.

    ``bar_enum_lab`` (A/B/IA/IB) permite diferenciar:
    - Tramos intermedios (no-\u00faltimo) con ``is_lap_extension_scheme``: +``embed`` al tronco.
    - Barras B/IB: el plano de corte est\u00e1 desplazado ``+embed`` respecto al de A, por lo que el
      tramo inferior de B es ``+embed`` m\u00e1s largo y el superior ``\u2212embed`` m\u00e1s corto.
    ``n_total_tramos`` = n\u00famero de tramos del troceo para determinar si es multi-tramo.
    ``top_collides`` = resultado de la prueba de colisi\u00f3n para el empotramiento superior:
    - ``True``  → colisi\u00f3n confirmada: el empotramiento se conserva (+embed al tronco).
    - ``False`` → sin colisi\u00f3n: empotramiento revertido, la barra no se extiende hacia arriba.
    - ``None``  → sin prueba (comportamiento anterior: +embed si is_top).
    """
    n_phys = len(seg_models or [])
    if not band_slot_indices or n_phys < 1:
        return u""
    try:
        from column_reinforcement_layout_rps import (
            LAYOUT_EMBED_CONCRETE_GRADE,
            column_bottom_joined_foundation_stretch_down_mm,
            _resolved_traslape_embed_mm,
            is_lap_extension_scheme,
            is_b_split_scheme,
        )
    except Exception:
        return u""

    grade_eff = concrete_grade
    if grade_eff is None:
        grade_eff = LAYOUT_EMBED_CONCRETE_GRADE

    traslape = _resolved_traslape_embed_mm(diameter_mm, grade_eff)
    if traslape is None:
        traslape = 0.0
    embed_mm = float(traslape)

    try:
        _lab = unicode(bar_enum_lab).strip() if bar_enum_lab is not None else u""
    except Exception:
        _lab = u""

    is_lap = is_lap_extension_scheme(_lab)
    is_b = is_b_split_scheme(_lab)

    n_tramos = max(1, int(n_total_tramos))
    is_multi = n_tramos > 1

    by_slot = {int(m[u"slot"]): m for m in (seg_models or [])}
    heights = []
    for sl in band_slot_indices:
        m = by_slot.get(int(sl))
        if m is None:
            continue
        try:
            heights.append(float(m[u"height_mm"]))
        except Exception:
            heights.append(0.0)
    trunk = sum(heights)
    # Empalmes entre columnas f\u00edsicas dentro de la misma banda
    if len(band_slot_indices) > 1:
        trunk += float(traslape) * float(len(band_slot_indices) - 1)

    has_bot = 0 in band_slot_indices
    has_top = (n_phys - 1) in band_slot_indices

    # Extensi\u00f3n inter-tramo (misma l\u00f3gica que seg_jobs: seg_i < n_seg_total − 1)
    if is_lap and is_multi and not has_top:
        trunk += embed_mm

    # Offset de plano B/IB: corte B = corte A + embed → tramo inferior de B m\u00e1s largo,
    # tramo superior de B m\u00e1s corto. En bandas que cubren toda la columna (has_bot y has_top)
    # el efecto se cancela y no se aplica.
    if is_b and is_multi and not (has_bot and has_top):
        if has_bot:
            trunk += embed_mm
        elif has_top:
            trunk -= embed_mm
    trunk = max(0.0, trunk)

    fund_mm = 0.0
    if has_bot and doc is not None:
        m0 = by_slot.get(0)
        if m0 is not None:
            el = m0.get(u"elem")
            if el is not None:
                try:
                    fund_mm = float(
                        column_bottom_joined_foundation_stretch_down_mm(doc, el),
                    )
                except Exception:
                    fund_mm = 0.0

    # Largo de pata horizontal (depende del Ø): aparece en el extremo inferior cuando hay
    # fundaci\u00f3n unida (want_bot_pata) y en el superior cuando se revert\u00eda el empotramiento
    # (want_top_pata).  Para el caso inferior lo calculamos directamente; para el superior
    # asumimos el camino nominal (empotramiento conservado, sin colisi\u00f3n) y NO sumamos pata.
    try:
        from column_reinforcement_layout_rps import _resolved_pata_hook_mm_for_revert
        pata_mm = float(_resolved_pata_hook_mm_for_revert(diameter_mm, grade_eff))
    except Exception:
        pata_mm = 0.0

    if has_bot:
        p_bot = float(fund_mm) if fund_mm > 1e-6 else embed_mm
    else:
        p_bot = 0.0
    # Empotramiento superior: si top_collides is False la barra no se extiende hacia arriba.
    # Si es True o None conservamos el comportamiento nominal (+embed al tronco superior).
    if has_top:
        p_top = 0.0 if (top_collides is False) else embed_mm
    else:
        p_top = 0.0
    # Pata inferior horizontal (solo cuando hay fundaci\u00f3n unida \u2192 want_bot_pata nominal)
    pata_bot_mm = pata_mm if (has_bot and fund_mm > 1e-6) else 0.0
    # Pata superior horizontal (solo cuando empotramiento revertido \u2192 want_top_pata)
    pata_top_mm = pata_mm if (has_top and top_collides is False) else 0.0

    total = trunk + p_bot + p_top + pata_bot_mm + pata_top_mm

    i_tot = int(round(total))
    # Componente vertical = tronco + extensi\u00f3n fundaci\u00f3n (sin patas horizontales)
    i_vert = int(round(trunk + p_bot)) if has_bot else int(round(trunk))
    it = int(round(p_top))      # empotramiento superior (barra entra en estructura)
    ip = int(round(pata_bot_mm))   # pata horizontal inicio (startpoint)
    itp = int(round(pata_top_mm))  # pata horizontal fin    (endpoint)

    # Orden canónico: (pata_start + vertical + embed_top|pata_end)
    # Solo se incluyen componentes > 0.
    parts = []
    if ip > 0:
        parts.append(ip)
    parts.append(i_vert)
    if it > 0:
        parts.append(it)
    if itp > 0:
        parts.append(itp)

    if len(parts) == 1:
        return u"L={0} ({1})".format(i_tot, parts[0])
    return u"L={0} ({1})".format(i_tot, u"+".join(str(p) for p in parts))


def _add_revit_level_head(canvas, line_y, bubble_cx, bubble_r, line_x1, line_x2, z_index=40):
    u"""L\u00ednea horizontal segmentada (``line_x1``..``line_x2``) y cabeza circular. La l\u00ednea usa Z bajo el fuste."""
    zi = int(z_index)
    ln = Line()
    ln.Stroke = _BR_LEVEL_LINE
    ln.StrokeThickness = 0.9
    try:
        dashes = DoubleCollection()
        dashes.Add(4.0)
        dashes.Add(3.0)
        ln.StrokeDashArray = dashes
    except Exception:
        pass
    x1 = float(line_x1)
    x2 = float(line_x2)
    ln.X1 = x1
    ln.Y1 = float(line_y)
    ln.X2 = max(x2, x1 + 1.0)
    ln.Y2 = float(line_y)
    try:
        Panel.SetZIndex(ln, int(_Z_TROCEO_DATUM_LINE))
    except Exception:
        pass
    canvas.Children.Add(ln)

    bx = float(bubble_cx) - float(bubble_r)
    by = float(line_y) - float(bubble_r)
    dia = 2.0 * float(bubble_r)

    disk = Ellipse()
    disk.Width = dia
    disk.Height = dia
    disk.Fill = _BR_LEVEL_DISK
    disk.Stroke = Brushes.Transparent
    disk.StrokeThickness = 0.95
    Canvas.SetLeft(disk, bx)
    Canvas.SetTop(disk, by)
    try:
        Panel.SetZIndex(disk, zi)
    except Exception:
        pass
    canvas.Children.Add(disk)
    zi += 1

    cx = float(bubble_cx)
    cy = float(line_y)
    r = float(bubble_r)
    try:
        # Ajedrezado: SO y NE claros (fondo disco), NO y SE en azul tema (cuartos desde el centro).
        p_tl = Path()
        p_tl.Data = Geometry.Parse(
            u"M {0},{1} L {2},{1} A {3},{3} 0 0 1 {0},{4} Z".format(
                cx,
                cy,
                cx - r,
                r,
                cy - r,
            )
        )
        p_tl.Fill = _BR_LEVEL_BUBBLE
        try:
            Panel.SetZIndex(p_tl, zi)
        except Exception:
            pass
        canvas.Children.Add(p_tl)
        zi += 1

        p_br = Path()
        p_br.Data = Geometry.Parse(
            u"M {0},{1} L {2},{1} A {3},{3} 0 0 1 {0},{4} Z".format(
                cx,
                cy,
                cx + r,
                r,
                cy + r,
            )
        )
        p_br.Fill = _BR_LEVEL_BUBBLE
        try:
            Panel.SetZIndex(p_br, zi)
        except Exception:
            pass
        canvas.Children.Add(p_br)
        zi += 1

        lead = max(4.0, 0.35 * r)
        ln_lead = Line()
        ln_lead.X1 = cx - r - lead
        ln_lead.X2 = cx - r
        ln_lead.Y1 = cy
        ln_lead.Y2 = cy
        ln_lead.Stroke = _BR_LEVEL_LINE
        ln_lead.StrokeThickness = max(0.85, 0.9)
        try:
            Panel.SetZIndex(ln_lead, zi)
        except Exception:
            pass
        canvas.Children.Add(ln_lead)
        zi += 1

        rim = Ellipse()
        rim.Width = dia
        rim.Height = dia
        rim.Fill = Brushes.Transparent
        rim.Stroke = _BR_LEVEL_TEXT
        rim.StrokeThickness = 0.95
        Canvas.SetLeft(rim, bx)
        Canvas.SetTop(rim, by)
        try:
            Panel.SetZIndex(rim, zi)
        except Exception:
            pass
        canvas.Children.Add(rim)
    except Exception:
        pass


class TroceoSchemeOutcome(object):
    """Resultado del diálogo de esquema."""

    def __init__(
        self,
        cancelled=False,
        skip_no_cut=False,
        columns=None,
        segment_rebar_bar_type_ids=None,
        troceo_empalme_policy_by_column_id=None,
    ):
        self.cancelled = bool(cancelled)
        self.skip_no_cut = bool(skip_no_cut)
        self.columns = list(columns) if columns else []
        self.segment_rebar_bar_type_ids = (
            [int(x) for x in segment_rebar_bar_type_ids]
            if segment_rebar_bar_type_ids is not None
            else None
        )
        _pol = {}
        for _k, _v in (troceo_empalme_policy_by_column_id or {}).items():
            try:
                _pol[int(_k)] = unicode(_v)
            except Exception:
                pass
        self.troceo_empalme_policy_by_column_id = _pol


def _load_xaml_text():
    here = os.path.dirname(os.path.abspath(__file__))
    xaml_path = os.path.join(here, "troceo_scheme_window.xaml")
    txt = File.ReadAllText(xaml_path)
    return txt.replace("__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)


class TroceoSchemeController(object):
    """Construye bloques apilados (alta → abajo en orden visual: superior = mayor Z)."""

    def __init__(
        self,
        rows,
        uiapp=None,
        uidoc=None,
        doc=None,
        default_bar_diam_mm=12.0,
        parent_window=None,
        blocks_host=None,
        diam_host=None,
        btn_confirm=None,
        btn_cancel=None,
        btn_alternate_sel=None,
        scheme_scrollviewer=None,
        diam_scrollviewer=None,
        embed_notify=None,
        tb_alternate_start=None,
        tb_alternate_step=None,
        column_stirrup_spacing_slot_store=None,
        column_stirrup_spacing_default_cb=None,
        column_stirrup_bar_type_slot_store=None,
        column_stirrup_bar_type_choices=None,
        column_stirrup_bar_type_default_cb=None,
        column_stirrup_policy_slot_store=None,
        longitudinal_line_ubicacion_labels=None,
        longitudinal_line_scheme_by_label=None,
    ):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        if self._doc is None and uidoc is not None:
            try:
                self._doc = uidoc.Document
            except Exception:
                self._doc = None
        self._default_diam_mm = float(default_bar_diam_mm)
        self._bar_choices = _collect_project_rebar_bar_types(self._doc)
        self._sel_slots = set()
        self._define_empalme_checkboxes = {}
        self._troceo_empalme_cb_suppress = False
        self._borders_by_slot = {}
        models = []
        for row in rows or []:
            elem, z_mm, eid, h_mm, lev, pilar_raw = _unpack_troceo_row(row)
            if eid < 0 or elem is None:
                continue
            lev_s = lev
            if lev_s is not None and hasattr(lev_s, "ToString"):
                try:
                    lev_s = lev_s.ToString()
                except Exception:
                    pass
            if lev_s is not None:
                try:
                    lev_s = unicode(lev_s)
                except Exception:
                    try:
                        lev_s = u"{0}".format(lev_s)
                    except Exception:
                        lev_s = u""
            pil_s = pilar_raw
            if pil_s is not None:
                try:
                    pil_s = unicode(pil_s)
                except Exception:
                    try:
                        pil_s = u"{0}".format(pil_s)
                    except Exception:
                        pil_s = None
                if pil_s is not None:
                    pil_s = pil_s.strip()
                    if not pil_s:
                        pil_s = None
            models.append(
                {
                    u"elem": elem,
                    u"z_mm": float(z_mm),
                    u"eid": int(eid),
                    u"height_mm": h_mm,
                    u"level_label": lev_s,
                    u"pilar_label": pil_s,
                }
            )
        models.sort(key=lambda p: p[u"z_mm"])
        for i, p in enumerate(models):
            p[u"slot"] = int(i)
        _finalize_segment_heights(models)
        self._seg_models = models
        self._row_entries = [
            (p[u"elem"], p[u"z_mm"], p[u"eid"], p[u"slot"]) for p in models
        ]
        self._pending_bar_type_ids = None
        self._diam_combos = []
        self._longitudinal_diam_type_ids = None
        self._scheme_bar_length_labels = []
        self._troceo_layout_scale = 1.0
        self._exact_band_lengths = {}  # (ubic_lab, tramo_i) → str
        self._diam_ladder_canvas = None
        self._longitudinal_line_ubicacion_labels = list(
            longitudinal_line_ubicacion_labels or [],
        )
        self._longitudinal_line_scheme_by_label = dict(
            longitudinal_line_scheme_by_label or {},
        )
        self._scheme_bar_length_label_matrix = []
        self._embed_notify = embed_notify
        self._scheme_scrollviewer = scheme_scrollviewer
        self._diam_scrollviewer = diam_scrollviewer
        self._scroll_bottom_timer = None
        self._diam_scroll_bottom_timer = None
        self._embedded = parent_window is not None
        self._tb_alternate_start = tb_alternate_start
        self._tb_alternate_step = tb_alternate_step
        self._col_spacing_store = column_stirrup_spacing_slot_store
        self._col_spacing_default_cb = column_stirrup_spacing_default_cb
        self._col_spacing_evt_suppress = False
        self._col_stirrup_bar_store = column_stirrup_bar_type_slot_store
        self._col_stirrup_bar_choices = list(column_stirrup_bar_type_choices or [])
        self._col_stirrup_bar_default_cb = column_stirrup_bar_type_default_cb
        self._col_bar_type_evt_suppress = False
        self._col_stirrup_policy_store = column_stirrup_policy_slot_store
        if self._col_stirrup_policy_store is None:
            self._col_stirrup_policy_store = {}
        self._col_policy_evt_suppress = False
        self._troceo_empalme_policy_by_slot = {}
        self._troceo_empalme_policy_combos = {}
        self._troceo_empalme_policy_evt_suppress = False
        if self._embedded:
            self.window = parent_window
            self._blocks_host = blocks_host
            self._diam_host = diam_host
            self._embed_btn_confirm = btn_confirm
            self._embed_btn_cancel = btn_cancel
            self._embed_btn_alternate_sel = btn_alternate_sel
        else:
            self._embed_btn_confirm = None
            self._embed_btn_cancel = None
            self._embed_btn_alternate_sel = None
            self.window = XamlReader.Parse(_load_xaml_text())
            self._blocks_host = self.window.FindName("BlocksHost")
            self._diam_host = self.window.FindName("DiamHost")
            if self._scheme_scrollviewer is None:
                try:
                    self._scheme_scrollviewer = self.window.FindName(
                        "TroceoSchemeScroll",
                    )
                except Exception:
                    self._scheme_scrollviewer = None
            if self._tb_alternate_start is None:
                try:
                    self._tb_alternate_start = self.window.FindName(
                        "TbTroceoAltStart",
                    )
                except Exception:
                    pass
            if self._tb_alternate_step is None:
                try:
                    self._tb_alternate_step = self.window.FindName(
                        "TbTroceoAltStep",
                    )
                except Exception:
                    pass
        if self._tb_alternate_start is None and self.window is not None:
            try:
                self._tb_alternate_start = self.window.FindName(
                    "TbTroceoAltStart",
                )
            except Exception:
                pass
        if self._tb_alternate_step is None and self.window is not None:
            try:
                self._tb_alternate_step = self.window.FindName("TbTroceoAltStep")
            except Exception:
                pass
        if self._diam_scrollviewer is None and self.window is not None:
            try:
                self._diam_scrollviewer = self.window.FindName(
                    "TroceoDiamScroll",
                )
            except Exception:
                self._diam_scrollviewer = None
        self._scheme_header_canvas = None
        if self.window is not None:
            try:
                self._scheme_header_canvas = self.window.FindName(
                    "TroceoSchemeHeaderCanvas",
                )
            except Exception:
                self._scheme_header_canvas = None
        self._blocks_reveal_timer = None
        self._blocks_reveal_batches = []
        self._blocks_reveal_i = 0
        self._viewport_stack_budget_px = None
        self._last_viewport_refit_h = None
        self._viewport_refit_running = False
        self._populate_blocks()
        self._wire_buttons()
        self._schedule_scroll_scheme_to_bottom()
        self._hook_embedded_viewport_refit()

    def _hook_embedded_viewport_refit(self):
        u"""Ajusta ``stack_h`` al alto útil del área fija del asistente (sin scrollbars)."""
        if not getattr(self, "_embedded", False):
            return
        sv = self._scheme_scrollviewer
        if sv is None:
            return
        try:
            from System import Action

            def run_refit():
                try:
                    self._refit_embedded_scheme_to_viewport(force=True)
                except Exception:
                    pass

            w = self.window
            if w is not None:
                try:
                    w.Dispatcher.BeginInvoke(
                        DispatcherPriority.ContextIdle,
                        Action(run_refit),
                    )
                except Exception:
                    try:
                        run_refit()
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            def on_size(sender, args):
                try:
                    self._refit_embedded_scheme_to_viewport(force=False)
                except Exception:
                    pass

            sv.SizeChanged += on_size
        except Exception:
            pass

    def _refit_embedded_scheme_to_viewport(self, force=False):
        if not getattr(self, "_embedded", False):
            return
        if getattr(self, "_viewport_refit_running", False):
            return
        sv = self._scheme_scrollviewer
        if sv is None:
            return
        try:
            h = float(sv.ActualHeight)
        except Exception:
            return
        if h < 60.0:
            return
        if not force and self._last_viewport_refit_h is not None:
            try:
                if abs(h - float(self._last_viewport_refit_h)) < 2.5:
                    return
            except Exception:
                pass
        self._last_viewport_refit_h = h
        ls = float(_TROCEO_LAYOUT_SCALE)
        # ``stack_h`` + ~``32*ls`` (cotas/cimentación) + márgenes del host ≈ alto contenido; no restar de más.
        overhead = 32.0 * ls + 12.0
        budget = max(80.0, h - overhead)
        prev = getattr(self, "_viewport_stack_budget_px", None)
        if (
            not force
            and prev is not None
            and abs(float(prev) - budget) < 1.5
        ):
            return
        self._viewport_stack_budget_px = budget
        self._viewport_refit_running = True
        try:
            self._populate_blocks(skip_reveal=True)
            try:
                self._refresh_all_shaft_styles()
            except Exception:
                pass
        finally:
            self._viewport_refit_running = False

    def _scroll_scheme_to_bottom(self):
        """Desplaza el ScrollViewer al final: base del pilar (troceo de abajo arriba)."""
        sv = getattr(self, "_scheme_scrollviewer", None)
        if sv is None:
            return
        try:
            h = float(sv.ScrollableHeight)
            if h > 0.5:
                sv.ScrollToVerticalOffset(h)
                return
        except Exception:
            pass
        try:
            sv.ScrollToEnd()
        except Exception:
            pass

    def _schedule_scroll_scheme_to_bottom(self):
        """Tras medir el contenido; ContextIdle + timer por animación/revelado."""
        sv = getattr(self, "_scheme_scrollviewer", None)
        if sv is None:
            return

        try:
            from System import Action

            self.window.Dispatcher.BeginInvoke(
                DispatcherPriority.ContextIdle,
                Action(self._scroll_scheme_to_bottom),
            )
        except Exception:
            try:
                self._scroll_scheme_to_bottom()
            except Exception:
                pass

        try:
            t_old = getattr(self, "_scroll_bottom_timer", None)
            if t_old is not None:
                try:
                    t_old.Stop()
                except Exception:
                    pass

            def on_tick(sender, args):
                try:
                    sender.Stop()
                except Exception:
                    pass
                self._scroll_scheme_to_bottom()

            t = DispatcherTimer()
            t.Interval = TimeSpan.FromMilliseconds(180.0)
            t.Tick += on_tick
            self._scroll_bottom_timer = t
            t.Start()
        except Exception:
            pass

    def _scroll_diam_to_bottom(self):
        """Scroll del panel Tramos al final (tramo inferior / base)."""
        sv = getattr(self, "_diam_scrollviewer", None)
        if sv is None:
            return
        try:
            h = float(sv.ScrollableHeight)
            if h > 0.5:
                sv.ScrollToVerticalOffset(h)
                return
        except Exception:
            pass
        try:
            sv.ScrollToEnd()
        except Exception:
            pass

    def _schedule_scroll_diam_to_bottom(self):
        sv = getattr(self, "_diam_scrollviewer", None)
        if sv is None:
            return
        try:
            from System import Action

            self.window.Dispatcher.BeginInvoke(
                DispatcherPriority.ContextIdle,
                Action(self._scroll_diam_to_bottom),
            )
        except Exception:
            try:
                self._scroll_diam_to_bottom()
            except Exception:
                pass
        try:
            t_old = getattr(self, "_diam_scroll_bottom_timer", None)
            if t_old is not None:
                try:
                    t_old.Stop()
                except Exception:
                    pass

            def on_tick(sender, args):
                try:
                    sender.Stop()
                except Exception:
                    pass
                self._scroll_diam_to_bottom()

            t = DispatcherTimer()
            t.Interval = TimeSpan.FromMilliseconds(180.0)
            t.Tick += on_tick
            self._diam_scroll_bottom_timer = t
            t.Start()
        except Exception:
            pass

    def _apply_troceo_combo_resources(self, cb):
        """Aplica plantilla oscura (popup + ítems) definida en BIMTOOLS_DARK_STYLES_XML."""
        w = self.window
        if w is None or cb is None:
            return False
        try:
            st = w.TryFindResource(u"ComboTroceoStretch")
            if st is None:
                st = w.TryFindResource(u"ComboStretch")
            if st is not None:
                cb.Style = st
                try:
                    cb.ClearValue(ComboBox.WidthProperty)
                    cb.ClearValue(ComboBox.HeightProperty)
                    cb.ClearValue(ComboBox.MaxWidthProperty)
                except Exception:
                    pass
                return True
        except Exception:
            pass
        try:
            st = w.TryFindResource(u"Combo")
            it = w.TryFindResource(u"ComboItem")
            if st is not None:
                cb.Style = st
                try:
                    cb.ClearValue(ComboBox.WidthProperty)
                    cb.ClearValue(ComboBox.MaxWidthProperty)
                except Exception:
                    pass
            if it is not None:
                cb.ItemContainerStyle = it
            return st is not None or it is not None
        except Exception:
            pass
        return False

    def _apply_scheme_spacing_combo_look(self, cb):
        u"""Estilo oscuro compacto: evita ComboStretch que ensancha y solapa el fuste."""
        w = self.window
        if w is None or cb is None:
            return
        try:
            st = w.TryFindResource(u"Combo")
            it = w.TryFindResource(u"ComboItem")
            if st is not None:
                cb.Style = st
            if it is not None:
                cb.ItemContainerStyle = it
        except Exception:
            pass
        try:
            cb.Background = _BR_ENTRY_BG
            cb.Foreground = _BR_ENTRY_FG
            cb.BorderBrush = _BR_ENTRY_BD
        except Exception:
            pass

    def _layout_bar_length_labels(self, canvas, bar_x0_ladder, bands, ls):
        u"""Etiquetas de largo en el canvas del bloque de tramos (bajo el esquema), a la izquierda de la escalerilla; texto vertical."""
        self._scheme_bar_length_labels = []
        self._scheme_bar_length_label_matrix = []
        if canvas is None or not bands:
            return
        _remove_bar_length_est_shapes(canvas)

    def _refresh_bar_tramos_ladder(self):
        u"""Limpia la escalerilla dibujada en el lienzo; los tramos se marcan en el fuste y el \u00d8 longitudinal va en el mismo canvas."""
        lay = getattr(self, "_bar_ladder_layout", None)
        cv_main = getattr(self, "_scheme_canvas", None)
        if lay is None:
            if cv_main is not None:
                _remove_bar_length_est_shapes(cv_main)
                _remove_bar_tramo_ladder_shapes(cv_main)
                _remove_troceo_shaft_tramo_labels(cv_main)
            self._scheme_bar_length_labels = []
            self._scheme_bar_length_label_matrix = []
            return
        if cv_main is not None:
            _remove_bar_tramo_ladder_shapes(cv_main)
            _remove_bar_length_est_shapes(cv_main)
        try:
            top_pad = float(lay[u"top_pad"])
            stack_h = float(lay[u"stack_h"])
            bx_main = float(lay[u"bar_x0"])
            slot_map = lay.get(u"slot_to_y_bottom") or {}
        except Exception:
            return
        try:
            sel = set(self._sel_slots) if self._sel_slots else set()
        except Exception:
            sel = set()
        bands = self._bar_tramo_bands_for_layout(sel, lay)
        try:
            ls = float(lay.get(u"layout_scale", 1.0))
        except Exception:
            ls = 1.0
        if cv_main is not None and bands:
            self._layout_bar_length_labels(cv_main, bx_main, bands, ls)
        elif cv_main is not None:
            self._scheme_bar_length_labels = []
            self._scheme_bar_length_label_matrix = []
        try:
            self._sync_shaft_tramo_band_labels()
        except Exception:
            pass

    def _troceo_diam_fixed_header_host(self):
        if self.window is None:
            return None
        try:
            return self.window.FindName("TroceoDiamHeaderHost")
        except Exception:
            return None

    def _clear_troceo_diam_fixed_headers(self):
        u"""Cabeceras antiguas del panel Tramos (XAML heredado); vac\u00edo si el host existe."""
        hh = self._troceo_diam_fixed_header_host()
        if hh is None:
            return
        try:
            hh.Children.Clear()
            hh.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def _place_longitudinal_diam_combos_on_scheme(self, segments):
        u"""Coloca combos por banda y etiqueta «Tramo n» (alineados con la fila f\u00edsica inferior de la banda si hay varias)."""
        cv = getattr(self, "_scheme_canvas", None)
        lay = getattr(self, "_bar_ladder_layout", None)
        if cv is None or lay is None or not segments:
            return
        try:
            slot_map = lay.get(u"slot_to_y_bottom") or {}
            top_pad = float(lay[u"top_pad"])
            stack_h = float(lay[u"stack_h"])
            bar_x0 = float(lay[u"bar_x0"])
            ls = float(lay.get(u"layout_scale", 1.0))
            sel = getattr(self, "_sel_slots", None) or set()
            bands = self._bar_tramo_bands_for_layout(sel, lay)
            trom_x0 = float(lay.get(u"tramo_label_x0", bar_x0 - 48.0))
            col_w = float(lay.get(u"tramo_label_col_w", 44.0))
            slot_y_span = lay.get(u"slot_to_y_span") or {}
        except Exception:
            return
        _remove_longitudinal_diam_scheme_widgets(cv)
        h_cb = float(_stirrup_combo_compact_height_px(ls))
        for y_top_cell, h_cell, tramo_no in bands:
            idx = int(tramo_no) - 1
            if idx < 0 or idx >= len(segments):
                continue
            cb = segments[idx]
            try:
                cb.Tag = _LONGITUDINAL_DIAM_COMBO_TAG
            except Exception:
                pass
            try:
                top = _troceo_long_combo_top_px_for_band(
                    y_top_cell, h_cell, slot_y_span, h_cb
                )
            except Exception:
                top = float(y_top_cell)
            Canvas.SetLeft(cb, float(bar_x0))
            Canvas.SetTop(cb, top)
            try:
                Panel.SetZIndex(cb, 9)
            except Exception:
                pass
            try:
                cv.Children.Add(cb)
            except Exception:
                pass
            tbt = TextBlock()
            tbt.Text = u"Tramo {0}".format(int(tramo_no))
            _apply_troceo_band_tramo_label_style(tbt, ls)
            try:
                tbt.Tag = _TROCEO_LONG_UI_DECOR_TAG
            except Exception:
                pass
            try:
                from System.Windows.Controls import ToolTip

                tip = ToolTip()
                tip.Content = (
                    u"Tramo {0} de troceo (toda la banda entre cortes). "
                    u"Con varias filas de piso, el selector se alinea con la fila inferior de la banda."
                ).format(int(tramo_no))
                tbt.ToolTip = tip
            except Exception:
                pass
            try:
                from System import Double
                from System.Windows import Size

                tbt.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
                tw = float(tbt.DesiredSize.Width)
                th = float(tbt.DesiredSize.Height)
            except Exception:
                tw, th = 36.0, 12.0
            try:
                lx = float(trom_x0) + max(0.0, (float(col_w) - tw) / 2.0)
                ly = float(top) + max((h_cb - th) / 2.0, 0.0)
                Canvas.SetLeft(tbt, lx)
                Canvas.SetTop(tbt, ly)
                Panel.SetZIndex(tbt, 9)
                cv.Children.Add(tbt)
            except Exception:
                pass

    def _refresh_diam_length_labels(self):
        u"""Actualiza texto de largo estimado (bloque de tramos bajo el esquema; vertical, izquierda de la escalerilla)."""
        matrix = getattr(self, "_scheme_bar_length_label_matrix", None) or []
        combos = getattr(self, "_diam_combos", None) or []
        lay = getattr(self, "_bar_ladder_layout", None)
        models = list(getattr(self, "_seg_models", None) or [])
        ubic_all = list(getattr(self, "_longitudinal_line_ubicacion_labels", None) or [])
        if not matrix or not combos:
            return
        if len(matrix) != len(combos):
            return
        slot_span = lay.get(u"slot_to_y_span") if lay else None
        if not lay or not slot_span:
            for row_tbs in matrix:
                for tb in row_tbs or []:
                    if tb is not None:
                        try:
                            tb.Text = u""
                        except Exception:
                            pass
                        _apply_bar_length_preview_foreground(tb, u"")
            return
        try:
            top_pad = float(lay[u"top_pad"])
            stack_h = float(lay[u"stack_h"])
            slot_map = lay.get(u"slot_to_y_bottom") or {}
        except Exception:
            for row_tbs in matrix:
                for tb in row_tbs or []:
                    if tb is not None:
                        try:
                            tb.Text = u""
                        except Exception:
                            pass
                        _apply_bar_length_preview_foreground(tb, u"")
            return
        try:
            sel = set(self._sel_slots) if self._sel_slots else set()
        except Exception:
            sel = set()
        bands = self._bar_tramo_bands_for_layout(sel, lay)
        concrete_grade = None
        try:
            ls_f = float(lay.get(u"layout_scale", 1.0))
        except Exception:
            ls_f = 1.0
        bx_lbl = float(lay[u"bar_x0"])
        # Pre-calculo de colisión para el empotramiento del techo de la pila.
        # Solo se ejecuta una vez (el resultado se cachea por Ø) para la banda que contiene
        # el tramo superior.  El cache se invalida si el usuario reinicia el wizard.
        n_phys_col = len(models)
        top_slot_idx = n_phys_col - 1 if n_phys_col > 0 else 0
        # Detectar qué banda contiene el slot más alto para precalcular colisión
        _top_band_idx = None
        for _bi, (y0_bi, h_bi, _tno_bi) in enumerate(bands):
            _slots_bi = _slots_overlapping_band_px(y0_bi, h_bi, slot_span)
            if top_slot_idx in _slots_bi:
                _top_band_idx = _bi
                break
        _top_band_collision = None
        if _top_band_idx is not None and _top_band_idx < len(combos):
            _top_d_mm = _nominal_mm_from_bar_combo(
                combos[_top_band_idx],
                self._bar_choices,
                self._default_diam_mm,
            )
            try:
                _top_band_collision = self._run_top_embed_collision_check(_top_d_mm, concrete_grade)
            except Exception:
                _top_band_collision = None
        for i, row_tbs in enumerate(matrix):
            if not row_tbs:
                continue
            if i >= len(combos) or i >= len(bands):
                for tb in row_tbs:
                    if tb is not None:
                        try:
                            tb.Text = u""
                        except Exception:
                            pass
                        _apply_bar_length_preview_foreground(tb, u"")
                continue
            try:
                y0, h_cell, _tno = bands[i]
                slots = _slots_overlapping_band_px(y0, h_cell, slot_span)
                if not slots:
                    for tb in row_tbs:
                        if tb is not None:
                            tb.Text = u""
                            _apply_bar_length_preview_foreground(tb, u"")
                    continue
                d_mm = _nominal_mm_from_bar_combo(
                    combos[i],
                    self._bar_choices,
                    self._default_diam_mm,
                )
                n_tramos = len(combos)
                for j, tb in enumerate(row_tbs):
                    if tb is None:
                        continue
                    try:
                        ubi = u""
                        if ubic_all and j < len(ubic_all):
                            try:
                                ubi = unicode(ubic_all[j]).strip()
                            except Exception:
                                ubi = u"{0}".format(ubic_all[j]).strip()
                        # Resolver el esquema Revit (A/B/IA/IB) para el cálculo de largos.
                        # El label de display (A/B/C/D) refleja Armadura_Ubicacion; el esquema
                        # internal se recupera del mapa scheme_by_label.
                        _scheme_map = getattr(self, "_longitudinal_line_scheme_by_label", None) or {}
                        scheme_for_length = _scheme_map.get(ubi, ubi) if ubi else ubi
                        band_top_collides = (
                            _top_band_collision if i == _top_band_idx else None
                        )
                        s = _format_long_bar_length_estimate_mm(
                            models,
                            slots,
                            d_mm,
                            concrete_grade,
                            doc=getattr(self, "_doc", None),
                            bar_enum_lab=scheme_for_length,
                            n_total_tramos=n_tramos,
                            top_collides=band_top_collides,
                        )
                        if not s:
                            tb.Text = u""
                            _apply_bar_length_preview_foreground(tb, u"")
                        else:
                            disp = u"{0} {1}".format(ubi, s) if ubi else s
                            tb.Text = disp
                            _apply_bar_length_preview_foreground(tb, disp)
                    except Exception:
                        try:
                            tb.Text = u""
                            _apply_bar_length_preview_foreground(tb, u"")
                        except Exception:
                            pass
                _apply_side_by_side_vertical_bar_length_geometry(
                    row_tbs,
                    y0,
                    h_cell,
                    bx_lbl,
                    ls_f,
                )
            except Exception:
                for tb in row_tbs:
                    try:
                        if tb is not None:
                            tb.Text = u""
                            _apply_bar_length_preview_foreground(tb, u"")
                    except Exception:
                        pass

    def _schedule_refresh_diam_length_labels(self):
        u"""Refresca largo presunto (empalmes/empotramiento tabular y fundaci\u00f3n si aplica)."""
        w = getattr(self, u"window", None)
        try:
            from System import Action
        except Exception:
            Action = None
        if w is None or Action is None:
            try:
                self._refresh_diam_length_labels()
            except Exception:
                pass
            return
        try:
            w.Dispatcher.BeginInvoke(
                DispatcherPriority.Input,
                Action(self._refresh_diam_length_labels),
            )
        except Exception:
            try:
                self._refresh_diam_length_labels()
            except Exception:
                pass

    def _on_longitudinal_diam_changed(self, sender, args):
        # Invalidar cache para recalcular con el nuevo Ø
        self._cached_top_collision = {}
        self._schedule_refresh_diam_length_labels()

    def _snapshot_longitudinal_diam_type_ids(self):
        u"""Guarda el \u00d8/tipo elegido por tramo para el siguiente repanel tras cambiar cortes."""
        out = []
        for cb in getattr(self, "_diam_combos", None) or []:
            try:
                sel = cb.SelectedItem
                out.append(int(sel.Tag) if sel is not None else None)
            except Exception:
                out.append(None)
        self._longitudinal_diam_type_ids = out

    def _sync_diameter_panel(self):
        u"""Rebuilt el panel de \u00d8 longitudinal: combos en el canvas del esquema (derecha del fuste)."""
        self._cached_top_collision = {}
        self._diam_ladder_canvas = None
        cv = getattr(self, "_scheme_canvas", None)
        host = getattr(self, "_diam_host", None)
        if host is not None:
            try:
                host.Children.Clear()
            except Exception:
                pass
        self._clear_troceo_diam_fixed_headers()
        if cv is not None:
            _remove_longitudinal_diam_scheme_widgets(cv)
        prev_combos = list(self._diam_combos) if self._diam_combos else []
        self._diam_combos = []
        n_ref = len(self._sel_slots)
        if not self._bar_choices:
            if cv is not None:
                _remove_bar_length_est_shapes(cv)
            self._scheme_bar_length_labels = []
            self._scheme_bar_length_label_matrix = []
            self._longitudinal_diam_type_ids = None
            lay = getattr(self, "_bar_ladder_layout", None)
            if cv is not None and lay is not None:
                try:
                    lx = float(lay[u"bar_x0"])
                    ty = float(lay[u"top_pad"]) + 4.0
                except Exception:
                    lx, ty = 12.0, 8.0
                err = TextBlock()
                err.Text = (
                    u"No hay RebarBarType en el proyecto. Cargue tipos de armadura antes de usar el troceo."
                )
                err.Foreground = _BR_MUTED
                err.TextWrapping = TextWrapping.Wrap
                err.MaxWidth = 300.0
                err.Tag = _TROCEO_BAR_TYPES_MSG_TAG
                Canvas.SetLeft(err, lx)
                Canvas.SetTop(err, ty)
                try:
                    Panel.SetZIndex(err, 25)
                except Exception:
                    pass
                try:
                    cv.Children.Add(err)
                except Exception:
                    pass
            return
        n_segs = max(1, n_ref + 1)
        raw_from_ui = []
        for cb in prev_combos:
            try:
                sel = cb.SelectedItem
                raw_from_ui.append(int(sel.Tag) if sel is not None else None)
            except Exception:
                raw_from_ui.append(None)
        if not raw_from_ui and getattr(self, "_longitudinal_diam_type_ids", None):
            raw_from_ui = list(self._longitudinal_diam_type_ids)
        prev_tags = _longitudinal_diam_tags_resized(raw_from_ui, n_segs)
        segments = []
        for i in range(n_segs):
            cb = ComboBox()
            combo_resolved = self._apply_troceo_combo_resources(cb)
            try:
                ls_tr = float(getattr(self, "_troceo_layout_scale", _TROCEO_LAYOUT_SCALE))
            except Exception:
                ls_tr = float(_TROCEO_LAYOUT_SCALE)
            w_tramo_cb = _stirrup_combo_row_width_px(ls_tr, True)
            h_tramo_cb = _stirrup_combo_compact_height_px(ls_tr)
            pv_tr = _stirrup_combo_vertical_padding_px(ls_tr)
            try:
                cb.MinWidth = 0.0
                cb.HorizontalAlignment = HorizontalAlignment.Left
                cb.Width = w_tramo_cb
                cb.MaxWidth = w_tramo_cb
                cb.MinHeight = h_tramo_cb
                cb.Height = h_tramo_cb
                cb.MaxHeight = h_tramo_cb
                cb.Padding = Thickness(3.0, pv_tr, 3.0, pv_tr)
            except Exception:
                pass
            if not combo_resolved:
                try:
                    cb.Background = _BR_ENTRY_BG
                    cb.Foreground = _BR_ENTRY_FG
                    cb.BorderBrush = _BR_ENTRY_BD
                except Exception:
                    pass
            for bid, label, _ds, tip in self._bar_choices:
                it = ComboBoxItem()
                it.Content = label
                it.Tag = int(bid)
                if tip:
                    try:
                        it.ToolTip = tip
                    except Exception:
                        pass
                if not combo_resolved:
                    try:
                        it.Background = _BR_ENTRY_BG
                        it.Foreground = _BR_ENTRY_FG
                    except Exception:
                        pass
                cb.Items.Add(it)
            pick_i = _default_combo_index_for_mm(
                self._bar_choices,
                self._default_diam_mm,
            )
            if i < len(prev_tags) and prev_tags[i] is not None:
                for j, ch in enumerate(self._bar_choices):
                    if int(ch[0]) == int(prev_tags[i]):
                        pick_i = j
                        break
            try:
                cb.SelectedIndex = int(pick_i)
            except Exception:
                pass
            if n_segs > 1:
                try:
                    cb.ToolTip = u"RebarBarType · Tramo {0} de {1}".format(i + 1, n_segs)
                except Exception:
                    pass
            segments.append(cb)
            self._diam_combos.append(cb)
        self._place_longitudinal_diam_combos_on_scheme(segments)
        try:
            from System import EventHandler
            from System.Windows import RoutedEventHandler

            h_sel = RoutedEventHandler(self._on_longitudinal_diam_changed)
            h_close = EventHandler(self._on_longitudinal_diam_changed)
            for cb in self._diam_combos:
                cb.SelectionChanged += h_sel
                try:
                    cb.DropDownClosed += h_close
                except Exception:
                    pass
        except Exception:
            try:
                for cb in self._diam_combos:
                    cb.SelectionChanged += self._on_longitudinal_diam_changed
                    try:
                        cb.DropDownClosed += self._on_longitudinal_diam_changed
                    except Exception:
                        pass
            except Exception:
                pass
        self._refresh_bar_tramos_ladder()
        self._refresh_diam_length_labels()
        self._snapshot_longitudinal_diam_type_ids()
        self._schedule_scroll_diam_to_bottom()

    def _stop_blocks_reveal_animation(self):
        t = getattr(self, "_blocks_reveal_timer", None)
        if t is not None:
            try:
                t.Stop()
            except Exception:
                pass
            self._blocks_reveal_timer = None

    def _finish_blocks_reveal_animation(self):
        self._stop_blocks_reveal_animation()
        for brd in self._borders_by_slot.values():
            try:
                brd.Opacity = 1.0
            except Exception:
                pass

    def _reveal_next_batch(self):
        batches = getattr(self, "_blocks_reveal_batches", None) or []
        i = getattr(self, "_blocks_reveal_i", 0)
        if i >= len(batches):
            self._stop_blocks_reveal_animation()
            return
        for brd in batches[i]:
            try:
                brd.Opacity = 1.0
            except Exception:
                pass
        self._blocks_reveal_i = i + 1
        if self._blocks_reveal_i >= len(batches):
            self._stop_blocks_reveal_animation()
            self._schedule_scroll_scheme_to_bottom()

    def _on_blocks_reveal_tick(self, sender, args):
        self._reveal_next_batch()

    def _start_blocks_reveal_animation(self, brd_list):
        """Revela bloques en lotes de bajo Z (abajo en lista) a alto Z."""
        self._stop_blocks_reveal_animation()
        if not brd_list:
            self._schedule_scroll_scheme_to_bottom()
            return
        order = list(reversed(brd_list))
        n = len(order)
        n_batches = min(8, max(2, (n + 2) // 3))
        batch_size = max(1, (n + n_batches - 1) // n_batches)
        batches = []
        for j in range(0, n, batch_size):
            batches.append(order[j : j + batch_size])
        self._blocks_reveal_batches = batches
        self._blocks_reveal_i = 0
        try:
            t = DispatcherTimer()
            t.Interval = TimeSpan.FromMilliseconds(68)
            t.Tick += self._on_blocks_reveal_tick
            self._blocks_reveal_timer = t
            self._reveal_next_batch()
            if self._blocks_reveal_i < len(self._blocks_reveal_batches):
                t.Start()
        except Exception:
            self._finish_blocks_reveal_animation()
            self._schedule_scroll_scheme_to_bottom()

    def _place_segment_stirrup_spacing_combo(
        self,
        canvas,
        combo_left_x,
        y_top,
        hpx,
        slot,
        eid,
        elem,
        combo_w_px,
        combo_h_px,
        layout_scale,
        store_key=None,
        band_y_top=None,
        band_hpx=None,
        lot_label=None,
    ):
        u"""Combo mm junto al tramo del fuste (``store_key`` = tramo×lote)."""
        if self._col_spacing_store is None:
            return
        if canvas is None:
            return
        try:
            from System.Windows.Controls import ComboBox, ToolTip
        except Exception:
            return
        ls = float(layout_scale)
        cb_w = max(float(combo_w_px), 20.0)
        cb_h = max(float(combo_h_px), 20.0)
        pv = _stirrup_combo_vertical_padding_px(ls)
        cb = ComboBox()
        self._apply_scheme_spacing_combo_look(cb)
        try:
            cb.MinWidth = 0.0
            cb.Width = cb_w
            cb.MaxWidth = cb_w
            cb.MinHeight = cb_h
            cb.Height = cb_h
            cb.MaxHeight = cb_h
            cb.FontSize = max(8.0, 8.25 * ls)
            cb.Padding = Thickness(3.0, pv, 3.0, pv)
            cb.VerticalAlignment = VerticalAlignment.Center
        except Exception:
            pass
        try:
            from System.Windows import TextTrimming

            cb.TextTrimming = TextTrimming.CharacterEllipsis
        except Exception:
            pass
        for sp in _STIRRUP_SPACING_MM_CHOICES:
            cb.Items.Add(sp)
        if store_key is None:
            store_key = _stirrup_slot_lot_store_key(slot, 0)
        mm_val = _stirrup_spacing_from_store(
            self._col_spacing_store,
            store_key,
            slot,
            elem,
            self._col_spacing_default_cb,
        )
        if mm_val is None:
            try:
                _sk = int(store_key)
                _lot_i = int(_sk) % 3
                if _lot_i in (0, 2):
                    mm_val = float(_STIRRUP_L3_T1_T3_DEFAULT_SPACING_MM)
                else:
                    mm_val = 200.0
            except Exception:
                mm_val = 200.0
        sp_str = str(int(round(mm_val)))
        pick = 4
        for i, it in enumerate(_STIRRUP_SPACING_MM_CHOICES):
            if it == sp_str:
                pick = i
                break
        self._col_spacing_evt_suppress = True
        try:
            cb.SelectedIndex = int(pick)
        except Exception:
            if cb.Items.Count > 0:
                cb.SelectedIndex = 4
        finally:
            self._col_spacing_evt_suppress = False
        try:
            tip = ToolTip()
            tip.Content = u"Espaciamiento estribos (mm) \u00b7 Tramo {0} \u00b7 Id {1}".format(
                int(slot),
                int(eid),
            )
            cb.ToolTip = tip
        except Exception:
            pass

        try:
            cb.Tag = int(store_key)
        except Exception:
            pass

        def _on_stirrup_spacing_sel(sender, args):
            if self._col_spacing_evt_suppress:
                return
            try:
                it = sender.SelectedItem
                if it is None:
                    return
                sk_i = int(sender.Tag)
                self._col_spacing_store[sk_i] = float(str(it))
            except Exception:
                pass

        cb.SelectionChanged += _on_stirrup_spacing_sel
        try:
            _by = float(band_y_top) if band_y_top is not None else float(y_top)
            _bh = float(band_hpx) if band_hpx is not None else float(hpx)
            top = _by + max((_bh - float(cb_h)) / 2.0, 0.0)
        except Exception:
            top = float(y_top)
        Canvas.SetLeft(cb, float(combo_left_x))
        Canvas.SetTop(cb, top)
        try:
            Panel.SetZIndex(cb, 8)
        except Exception:
            pass
        try:
            canvas.Children.Add(cb)
        except Exception:
            pass

    def _place_segment_define_empalme_checkbox(
        self,
        canvas,
        empalme_x,
        empalme_max_w,
        y_top,
        hpx,
        slot,
        layout_scale,
    ):
        u"""Checkbox «Define Empalme» junto al fuste; el combo Ø longitudinal va a su derecha."""
        if canvas is None:
            return
        try:
            from System.Windows.Controls import CheckBox, ToolTip
        except Exception:
            return
        ls = float(layout_scale)
        chk = CheckBox()
        chk.Tag = int(slot)
        try:
            chk.Content = u"Define Empalme"
            chk.Foreground = _BR_TEXT
            chk.FontSize = max(7.0, 7.35 * ls)
            mw = max(52.0, float(empalme_max_w))
            chk.MaxWidth = mw
            chk.VerticalAlignment = VerticalAlignment.Center
        except Exception:
            pass
        try:
            self._troceo_empalme_cb_suppress = True
            chk.IsChecked = int(slot) in self._sel_slots
        except Exception:
            pass
        finally:
            self._troceo_empalme_cb_suppress = False
        try:
            tip = ToolTip()
            tip.Content = (
                u"Marca esta columna como referencia de plano para troceo y empalme."
            )
            chk.ToolTip = tip
        except Exception:
            pass
        try:
            from System import Double
            from System.Windows import Size
            from System.Windows import RoutedEventHandler

            def _emp_handler(sender, args):
                self._on_define_empalme_checkbox_changed(sender, args)

            chk.Checked += RoutedEventHandler(_emp_handler)
            chk.Unchecked += RoutedEventHandler(_emp_handler)
        except Exception:
            chk.Checked += self._on_define_empalme_checkbox_changed
            chk.Unchecked += self._on_define_empalme_checkbox_changed
        try:
            chk.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            ch = float(chk.DesiredSize.Height)
        except Exception:
            ch = max(22.0, 20.0 * ls)
        try:
            top = float(y_top) + max((float(hpx) - ch) / 2.0, 0.0)
        except Exception:
            top = float(y_top)
        Canvas.SetLeft(chk, float(empalme_x))
        Canvas.SetTop(chk, top)
        try:
            Panel.SetZIndex(chk, 15)
        except Exception:
            pass
        try:
            canvas.Children.Add(chk)
        except Exception:
            pass
        try:
            self._define_empalme_checkboxes[int(slot)] = chk
        except Exception:
            pass

    def _troceo_empalme_policy_for_slot(self, slot):
        try:
            p = self._troceo_empalme_policy_by_slot.get(int(slot))
            if p == TROCEO_EMPALME_POLICY_MID_AXIS:
                return TROCEO_EMPALME_POLICY_MID_AXIS
        except Exception:
            pass
        return TROCEO_EMPALME_POLICY_BASE

    def _sync_stirrup_policy_with_troceo_empalme(self, slot):
        u"""Mitad altura (empalme) \u2192 estribos L/3; Base \u2192 Completo (solo con Define Empalme)."""
        if self._col_stirrup_policy_store is None:
            return False
        try:
            slot_i = int(slot)
        except Exception:
            return False
        if slot_i not in self._sel_slots:
            return False
        troceo_pol = self._troceo_empalme_policy_for_slot(slot_i)
        if troceo_pol == TROCEO_EMPALME_POLICY_MID_AXIS:
            new_st = STIRRUP_POLICY_THIRDS_L3
        else:
            new_st = STIRRUP_POLICY_CONTINUOUS
        old_st = _stirrup_policy_for_slot(self._col_stirrup_policy_store, slot_i)
        if old_st == new_st:
            return False
        self._col_stirrup_policy_store[slot_i] = new_st
        if new_st == STIRRUP_POLICY_THIRDS_L3:
            self._init_lot_stores_from_slot(slot_i)
        try:
            self._populate_blocks(skip_reveal=True)
        except Exception:
            pass
        try:
            self._update_troceo_datum_reference_lines()
            self._refresh_all_shaft_styles()
            self._sync_shaft_tramo_band_labels()
            self._refresh_bar_tramos_ladder()
        except Exception:
            pass
        return True

    def _sync_troceo_empalme_policy_combo_state(self, slot):
        cb = (self._troceo_empalme_policy_combos or {}).get(int(slot))
        if cb is None:
            return
        try:
            active = int(slot) in self._sel_slots
            cb.IsEnabled = bool(active)
            cb.Opacity = 1.0 if active else 0.45
        except Exception:
            pass

    def _sync_all_troceo_empalme_policy_combos(self):
        for slot in (self._troceo_empalme_policy_combos or {}).keys():
            self._sync_troceo_empalme_policy_combo_state(slot)

    def _place_troceo_empalme_policy_combo(
        self,
        canvas,
        policy_x,
        y_top,
        hpx,
        slot,
        policy_w_px,
        combo_h_px,
        layout_scale,
    ):
        u"""Combo Base / Mitad altura; solo activo si «Define Empalme» está marcado."""
        if canvas is None:
            return
        try:
            from System.Windows.Controls import ComboBox, ToolTip
        except Exception:
            return
        ls = float(layout_scale)
        cb_w = max(float(policy_w_px), 80.0)
        cb_h = max(float(combo_h_px), 20.0)
        pv = _stirrup_combo_vertical_padding_px(ls)
        cb = ComboBox()
        self._apply_scheme_spacing_combo_look(cb)
        try:
            cb.Width = cb_w
            cb.MaxWidth = cb_w
            cb.MinHeight = cb_h
            cb.Height = cb_h
            cb.FontSize = max(7.5, 8.0 * ls)
            cb.Padding = Thickness(3.0, pv, 3.0, pv)
        except Exception:
            pass
        for _lbl, _val in _TROCEO_EMPALME_POLICY_UI_CHOICES:
            cb.Items.Add(_lbl)
        pol = self._troceo_empalme_policy_for_slot(slot)
        pick = 1 if pol == TROCEO_EMPALME_POLICY_MID_AXIS else 0
        self._troceo_empalme_policy_evt_suppress = True
        try:
            cb.SelectedIndex = int(pick)
        except Exception:
            pass
        finally:
            self._troceo_empalme_policy_evt_suppress = False
        try:
            if int(slot) not in self._troceo_empalme_policy_by_slot:
                self._troceo_empalme_policy_by_slot[int(slot)] = pol
        except Exception:
            pass
        try:
            tip = ToolTip()
            tip.Content = (
                u"Base: troceo en inicio del eje; estribos Completo. "
                u"Mitad altura: troceo al 50 % del sólido; estribos L/3 (T1\u2013T3); "
                u"traslapo 50/50 en la junta."
            )
            cb.ToolTip = tip
        except Exception:
            pass
        cb.Tag = int(slot)

        def _on_pol_sel(sender, args):
            if self._troceo_empalme_policy_evt_suppress:
                return
            try:
                slot_i = int(sender.Tag)
                ix = int(sender.SelectedIndex)
                if 0 <= ix < len(_TROCEO_EMPALME_POLICY_UI_CHOICES):
                    self._troceo_empalme_policy_by_slot[slot_i] = (
                        _TROCEO_EMPALME_POLICY_UI_CHOICES[ix][1]
                    )
            except Exception:
                pass
            try:
                if self._sync_stirrup_policy_with_troceo_empalme(slot_i):
                    return
            except Exception:
                pass
            try:
                self._update_troceo_datum_reference_lines()
                self._refresh_all_shaft_styles()
                self._sync_shaft_tramo_band_labels()
                self._refresh_bar_tramos_ladder()
            except Exception:
                pass

        cb.SelectionChanged += _on_pol_sel
        try:
            top = float(y_top) + max((float(hpx) - float(cb_h)) / 2.0, 0.0)
        except Exception:
            top = float(y_top)
        Canvas.SetLeft(cb, float(policy_x))
        Canvas.SetTop(cb, top)
        try:
            Panel.SetZIndex(cb, 14)
        except Exception:
            pass
        canvas.Children.Add(cb)
        self._troceo_empalme_policy_combos[int(slot)] = cb
        self._sync_troceo_empalme_policy_combo_state(int(slot))

    def _collect_troceo_empalme_policy_by_column_id(self):
        u"""Política troceo por id de instancia Revit (solo filas con empalme)."""
        try:
            from column_reinforcement_layout_rps import _element_id_iv
        except Exception:
            _element_id_iv = None
        out = {}
        for _elem, _z, eid, slot in getattr(self, "_row_entries", None) or []:
            try:
                s = int(slot)
            except Exception:
                continue
            if s not in self._sel_slots:
                continue
            pol = self._troceo_empalme_policy_for_slot(s)
            k = None
            if _element_id_iv is not None and _elem is not None:
                try:
                    k = int(_element_id_iv(_elem))
                except Exception:
                    k = None
            if k is None or k < 0:
                try:
                    k = int(eid)
                except Exception:
                    continue
            if k >= 0:
                out[k] = pol
        return out

    def _adjust_define_empalme_checkboxes_vs_long_combo(
        self, seg_layout, slot_to_y_bottom, top_pad, stack_h, layout_scale
    ):
        u"""Evita solape vertical casilla/combo salvo en la fila ancla, alineada con el centro del combo."""
        try:
            from System import Double
            from System.Windows import Size
        except Exception:
            return
        cb_map = getattr(self, "_define_empalme_checkboxes", None) or {}
        if not cb_map or not seg_layout:
            return
        ls = float(layout_scale)
        try:
            sel = set(self._sel_slots) if self._sel_slots else set()
            slot_span = {}
            for y_t, hp, sl, _ in seg_layout:
                try:
                    y0 = float(y_t)
                    slot_span[int(sl)] = (y0, y0 + float(hp))
                except Exception:
                    pass
            cut_map = {}
            for sl in sel:
                yc = self._troceo_datum_y_px_for_slot(
                    sl, slot_to_y_bottom or {}, slot_span
                )
                if yc is not None:
                    cut_map[int(sl)] = float(yc)
            bands = _bar_tramo_y_bands_from_cuts(
                sel,
                slot_to_y_bottom or {},
                float(top_pad),
                float(stack_h),
                slot_to_cut_y=cut_map,
            )
        except Exception:
            return
        if not bands:
            return
        h_cb = float(_stirrup_combo_compact_height_px(ls))

        def _overlaps(cya, cyb, t0, t1):
            return (min(cyb, t1) - max(cya, t0)) > 0.5

        def _resolve_top(y0s, ys1, chh, cy0, cy1):
            top_v = y0s + max((ys1 - y0s - chh) / 2.0, 0.0)
            if cy0 is None or cy1 is None:
                return top_v
            if not _overlaps(cy0, cy1, top_v, top_v + chh):
                return top_v
            margin = max(3.0, 3.0 * ls)
            t_down = min(max(top_v, cy1 + margin), ys1 - chh)
            if not _overlaps(cy0, cy1, t_down, t_down + chh):
                return t_down
            t_bot = max(y0s, ys1 - chh)
            if not _overlaps(cy0, cy1, t_bot, t_bot + chh):
                return t_bot
            t_hi = y0s
            if not _overlaps(cy0, cy1, t_hi, t_hi + chh):
                return t_hi
            return top_v

        for y_t, hp, sl, _ in seg_layout:
            sl_i = int(sl)
            chk = cb_map.get(sl_i)
            if chk is None:
                continue
            span = slot_span.get(sl_i)
            if not span:
                continue
            y0s, ys1 = float(span[0]), float(span[1])
            y_mid = 0.5 * (y0s + ys1)
            cy0 = cy1 = None
            band_ht = None
            for y_top_cell, h_cell, _tno in bands:
                yc0 = float(y_top_cell)
                yc1 = yc0 + float(h_cell)
                if (yc0 - 1e-6) <= y_mid <= (yc1 + 1e-6):
                    tcombo = _troceo_long_combo_top_px_for_band(
                        y_top_cell, h_cell, slot_span, h_cb
                    )
                    cy0, cy1 = tcombo, tcombo + h_cb
                    band_ht = (y_top_cell, h_cell)
                    break
            if cy0 is None:
                continue
            try:
                chk.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
                chh = float(chk.DesiredSize.Height)
            except Exception:
                chh = max(22.0, 20.0 * ls)
            anchor_sl = None
            if band_ht is not None:
                anchor_sl = _troceo_anchor_slot_index_for_band(
                    band_ht[0], band_ht[1], slot_span
                )
            try:
                if anchor_sl is not None and int(sl_i) == int(anchor_sl):
                    combo_cy = float(cy0) + float(h_cb) / 2.0
                    top_align = combo_cy - chh / 2.0
                    top_align = max(y0s, min(float(top_align), ys1 - chh))
                    Canvas.SetTop(chk, top_align)
                else:
                    Canvas.SetTop(
                        chk, _resolve_top(y0s, ys1, chh, cy0, cy1)
                    )
            except Exception:
                pass

    def _place_segment_stirrup_bar_type_combo(
        self,
        canvas,
        combo_left_x,
        y_top,
        hpx,
        slot,
        eid,
        elem,
        combo_w_px,
        combo_h_px,
        layout_scale,
        store_key=None,
        band_y_top=None,
        band_hpx=None,
    ):
        u"""Combo tipo/di\u00e1metro de estribo por tramo/lote (``store_key``)."""
        if self._col_stirrup_bar_store is None:
            return
        if not self._col_stirrup_bar_choices:
            return
        if canvas is None:
            return
        try:
            from System.Windows.Controls import ComboBox, ToolTip
        except Exception:
            return
        ls = float(layout_scale)
        cb_w = max(float(combo_w_px), 20.0)
        cb_h = max(float(combo_h_px), 20.0)
        pv = _stirrup_combo_vertical_padding_px(ls)
        cb = ComboBox()
        self._apply_scheme_spacing_combo_look(cb)
        try:
            cb.MinWidth = 0.0
            cb.Width = cb_w
            cb.MaxWidth = cb_w
            cb.MinHeight = cb_h
            cb.Height = cb_h
            cb.MaxHeight = cb_h
            cb.FontSize = max(8.0, 8.25 * ls)
            cb.Padding = Thickness(3.0, pv, 3.0, pv)
            cb.VerticalAlignment = VerticalAlignment.Center
        except Exception:
            pass
        try:
            from System.Windows import TextTrimming

            cb.TextTrimming = TextTrimming.CharacterEllipsis
        except Exception:
            pass
        for _lb, _bt in self._col_stirrup_bar_choices:
            try:
                cb.Items.Add(_lb)
            except Exception:
                cb.Items.Add(unicode(_lb))
        if store_key is None:
            store_key = _stirrup_slot_lot_store_key(slot, 0)
        bt_target = _stirrup_bar_type_from_store(
            self._col_stirrup_bar_store,
            store_key,
            slot,
            elem,
            self._col_stirrup_bar_default_cb,
            self._col_stirrup_bar_choices,
        )
        pick = 0
        _stirrup_pick_matched = False
        if bt_target is not None:
            _tid_tar = _rebar_type_id_int(bt_target)
            for i, (_lb, _bt) in enumerate(self._col_stirrup_bar_choices):
                try:
                    if (
                        _tid_tar is not None
                        and _rebar_type_id_int(_bt) == _tid_tar
                    ):
                        pick = i
                        _stirrup_pick_matched = True
                        break
                except Exception:
                    if _bt == bt_target:
                        pick = i
                        _stirrup_pick_matched = True
                        break
        self._col_bar_type_evt_suppress = True
        try:
            cb.SelectedIndex = int(pick)
        except Exception:
            if cb.Items.Count > 0:
                cb.SelectedIndex = 0
        finally:
            self._col_bar_type_evt_suppress = False
        try:
            if 0 <= int(pick) < len(self._col_stirrup_bar_choices):
                _chosen_bt = self._col_stirrup_bar_choices[int(pick)][1]
                if bt_target is None:
                    self._col_stirrup_bar_store[int(store_key)] = _chosen_bt
                elif _stirrup_pick_matched:
                    self._col_stirrup_bar_store[int(store_key)] = _chosen_bt
                else:
                    try:
                        if int(store_key) not in self._col_stirrup_bar_store:
                            self._col_stirrup_bar_store[int(store_key)] = bt_target
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            tip = ToolTip()
            tip.Content = u"Tipo / \u00d8 estribos \u00b7 Id {0}".format(int(eid))
            cb.ToolTip = tip
        except Exception:
            pass
        try:
            cb.Tag = int(store_key)
        except Exception:
            pass

        def _on_stirrup_bar_type_sel(sender, args):
            if self._col_bar_type_evt_suppress:
                return
            try:
                sk_i = int(sender.Tag)
                ix = int(sender.SelectedIndex)
                if 0 <= ix < len(self._col_stirrup_bar_choices):
                    self._col_stirrup_bar_store[sk_i] = self._col_stirrup_bar_choices[ix][
                        1
                    ]
            except Exception:
                pass

        cb.SelectionChanged += _on_stirrup_bar_type_sel
        try:
            _by = float(band_y_top) if band_y_top is not None else float(y_top)
            _bh = float(band_hpx) if band_hpx is not None else float(hpx)
            top = _by + max((_bh - float(cb_h)) / 2.0, 0.0)
        except Exception:
            top = float(y_top)
        Canvas.SetLeft(cb, float(combo_left_x))
        Canvas.SetTop(cb, top)
        try:
            Panel.SetZIndex(cb, 9)
        except Exception:
            pass
        try:
            canvas.Children.Add(cb)
        except Exception:
            pass

    def _init_lot_stores_from_slot(self, slot):
        u"""Al activar L/3: T1/T3 esp. 100 mm; T2 copia Ø/esp del lote único si faltan."""
        sk0 = _stirrup_slot_lot_store_key(slot, 0)
        for lot_i in range(3):
            sk = _stirrup_slot_lot_store_key(slot, lot_i)
            if self._col_spacing_store is not None:
                try:
                    if lot_i in (0, 2):
                        self._col_spacing_store[sk] = float(
                            _STIRRUP_L3_T1_T3_DEFAULT_SPACING_MM
                        )
                    elif sk not in self._col_spacing_store:
                        if sk0 in self._col_spacing_store:
                            self._col_spacing_store[sk] = float(
                                self._col_spacing_store[sk0]
                            )
                        elif int(slot) in self._col_spacing_store:
                            self._col_spacing_store[sk] = float(
                                self._col_spacing_store[int(slot)]
                            )
                except Exception:
                    pass
            if self._col_stirrup_bar_store is not None:
                try:
                    if sk not in self._col_stirrup_bar_store:
                        if sk0 in self._col_stirrup_bar_store:
                            self._col_stirrup_bar_store[sk] = self._col_stirrup_bar_store[
                                sk0
                            ]
                        elif int(slot) in self._col_stirrup_bar_store:
                            self._col_stirrup_bar_store[sk] = self._col_stirrup_bar_store[
                                int(slot)
                            ]
                except Exception:
                    pass

    def _draw_stirrup_third_guides(self, canvas, shaft_x, y_top, hpx, shaft_w, ls):
        u"""Líneas horizontales T1|T2|T3 en el fuste (política L/3)."""
        if canvas is None or float(hpx) < 6.0:
            return
        try:
            from System.Windows.Shapes import Line
            from System.Windows.Media import SolidColorBrush
        except Exception:
            return
        try:
            brush = _BR_MUTED
        except Exception:
            return
        third = float(hpx) / 3.0
        x0 = float(shaft_x)
        x1 = x0 + float(shaft_w)
        for k in (1, 2):
            try:
                y_line = float(y_top) + third * float(k)
                ln = Line()
                ln.X1 = x0
                ln.X2 = x1
                ln.Y1 = y_line
                ln.Y2 = y_line
                ln.Stroke = brush
                ln.StrokeThickness = max(0.75, 0.85 * float(ls))
                ln.StrokeDashArray = [3.0, 2.0]
                try:
                    Panel.SetZIndex(ln, 11)
                except Exception:
                    pass
                canvas.Children.Add(ln)
            except Exception:
                pass

    def _place_stirrup_lot_label(
        self,
        canvas,
        lote_x,
        y_top,
        hpx,
        lot_label,
        layout_scale,
        band_y_top=None,
        band_hpx=None,
    ):
        if canvas is None or not lot_label:
            return
        try:
            from System import Double
            from System.Windows import Size
            from System.Windows.Controls import TextBlock
        except Exception:
            return
        ls = float(layout_scale)
        tb = TextBlock()
        tb.Text = unicode(lot_label)
        try:
            tb.Foreground = _BR_MUTED
            tb.FontSize = max(8.0, 8.5 * ls)
        except Exception:
            pass
        try:
            _by = float(band_y_top) if band_y_top is not None else float(y_top)
            _bh = float(band_hpx) if band_hpx is not None else float(hpx)
            tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            lh = float(tb.DesiredSize.Height)
            top = _by + max((_bh - lh) / 2.0, 0.0)
            lw = float(tb.DesiredSize.Width)
            Canvas.SetLeft(
                tb,
                float(lote_x)
                + max(0.0, (_stirrup_lote_label_col_width_px(ls) - lw) / 2.0),
            )
            Canvas.SetTop(tb, top)
            Panel.SetZIndex(tb, 8)
            canvas.Children.Add(tb)
        except Exception:
            pass

    def _place_segment_stirrup_policy_combo(
        self,
        canvas,
        policy_x,
        y_top,
        hpx,
        slot,
        eid,
        policy_w_px,
        combo_h_px,
        layout_scale,
    ):
        if self._col_stirrup_policy_store is None or canvas is None:
            return
        try:
            from System.Windows.Controls import ComboBox, ToolTip
        except Exception:
            return
        ls = float(layout_scale)
        cb_w = max(float(policy_w_px), 36.0)
        cb_h = max(float(combo_h_px), 20.0)
        pv = _stirrup_combo_vertical_padding_px(ls)
        cb = ComboBox()
        self._apply_scheme_spacing_combo_look(cb)
        try:
            cb.MinWidth = cb_w
            cb.Width = cb_w
            cb.MaxWidth = cb_w
            cb.Height = cb_h
            cb.MaxHeight = cb_h
            cb.FontSize = max(8.0, 8.25 * ls)
            cb.Padding = Thickness(3.0, pv, 3.0, pv)
            cb.HorizontalAlignment = HorizontalAlignment.Left
        except Exception:
            pass
        for _lbl, _val in _STIRRUP_POLICY_UI_CHOICES:
            cb.Items.Add(_lbl)
        pol = _stirrup_policy_for_slot(self._col_stirrup_policy_store, slot)
        pick = 0 if pol == STIRRUP_POLICY_CONTINUOUS else 1
        self._col_policy_evt_suppress = True
        try:
            cb.SelectedIndex = int(pick)
        except Exception:
            pass
        finally:
            self._col_policy_evt_suppress = False
        try:
            if int(slot) not in self._col_stirrup_policy_store:
                self._col_stirrup_policy_store[int(slot)] = pol
        except Exception:
            pass
        try:
            tip = ToolTip()
            tip.Content = (
                u"Completo: un \u00d8 y espaciamiento en toda la columna. "
                u"L/3: tres lotes (T1\u2013T3) con \u00d8 y esp. independientes."
            )
            cb.ToolTip = tip
        except Exception:
            pass
        cb.Tag = int(slot)

        def _on_policy_sel(sender, args):
            if self._col_policy_evt_suppress:
                return
            try:
                slot_i = int(sender.Tag)
                ix = int(sender.SelectedIndex)
                if ix < 0 or ix >= len(_STIRRUP_POLICY_UI_CHOICES):
                    return
                new_pol = _STIRRUP_POLICY_UI_CHOICES[ix][1]
                old_pol = _stirrup_policy_for_slot(
                    self._col_stirrup_policy_store,
                    slot_i,
                )
                self._col_stirrup_policy_store[slot_i] = new_pol
                if (
                    new_pol == STIRRUP_POLICY_THIRDS_L3
                    and old_pol != STIRRUP_POLICY_THIRDS_L3
                ):
                    self._init_lot_stores_from_slot(slot_i)
                self._populate_blocks(skip_reveal=True)
            except Exception:
                pass

        cb.SelectionChanged += _on_policy_sel
        try:
            top = float(y_top) + max((float(hpx) - float(cb_h)) / 2.0, 0.0)
        except Exception:
            top = float(y_top)
        Canvas.SetLeft(cb, float(policy_x))
        Canvas.SetTop(cb, top)
        try:
            Panel.SetZIndex(cb, 9)
        except Exception:
            pass
        canvas.Children.Add(cb)

    def _place_column_stirrup_row_controls(
        self,
        canvas,
        lote_x,
        combo_left_x,
        policy_x,
        y_top,
        hpx,
        slot,
        eid,
        elem,
        stirrup_diam_w_px,
        stirrup_spacing_w_px,
        stirrup_combo_h_px,
        policy_w_px,
        ls,
        _has_stirrup_bar_combo,
        stirrup_combo_inner_gap_px,
        shaft_x,
        shaft_w,
    ):
        u"""Ø, esp. y política por columna; tres lotes si L/3."""
        if self._col_spacing_store is None:
            return
        pol = _stirrup_policy_for_slot(self._col_stirrup_policy_store, slot)
        _spacing_x_base = float(combo_left_x)
        if _has_stirrup_bar_combo and stirrup_diam_w_px > 0.5:
            _diam_x = float(combo_left_x)
        else:
            _diam_x = _spacing_x_base
        if pol == STIRRUP_POLICY_THIRDS_L3:
            self._draw_stirrup_third_guides(
                canvas, shaft_x, y_top, hpx, shaft_w, ls
            )
            third_h = float(hpx) / 3.0
            for lot_i in range(3):
                band_y = float(y_top) + third_h * float(lot_i)
                sk = _stirrup_slot_lot_store_key(slot, lot_i)
                self._place_stirrup_lot_label(
                    canvas,
                    lote_x,
                    y_top,
                    hpx,
                    _STIRRUP_LOT_LABELS[lot_i],
                    ls,
                    band_y_top=band_y,
                    band_hpx=third_h,
                )
                if _has_stirrup_bar_combo and stirrup_diam_w_px > 0.5:
                    self._place_segment_stirrup_bar_type_combo(
                        canvas,
                        _diam_x,
                        y_top,
                        hpx,
                        int(slot),
                        eid,
                        elem,
                        stirrup_diam_w_px,
                        stirrup_combo_h_px,
                        ls,
                        store_key=sk,
                        band_y_top=band_y,
                        band_hpx=third_h,
                    )
                    _sp_x = (
                        float(_diam_x)
                        + float(stirrup_diam_w_px)
                        + float(stirrup_combo_inner_gap_px)
                    )
                else:
                    _sp_x = _spacing_x_base
                self._place_segment_stirrup_spacing_combo(
                    canvas,
                    _sp_x,
                    y_top,
                    hpx,
                    int(slot),
                    eid,
                    elem,
                    stirrup_spacing_w_px,
                    stirrup_combo_h_px,
                    ls,
                    store_key=sk,
                    band_y_top=band_y,
                    band_hpx=third_h,
                )
        else:
            sk = _stirrup_slot_lot_store_key(slot, 0)
            if _has_stirrup_bar_combo and stirrup_diam_w_px > 0.5:
                self._place_segment_stirrup_bar_type_combo(
                    canvas,
                    _diam_x,
                    y_top,
                    hpx,
                    int(slot),
                    eid,
                    elem,
                    stirrup_diam_w_px,
                    stirrup_combo_h_px,
                    ls,
                    store_key=sk,
                )
                _sp_x = (
                    float(_diam_x)
                    + float(stirrup_diam_w_px)
                    + float(stirrup_combo_inner_gap_px)
                )
            else:
                _sp_x = _spacing_x_base
            self._place_segment_stirrup_spacing_combo(
                canvas,
                _sp_x,
                y_top,
                hpx,
                int(slot),
                eid,
                elem,
                stirrup_spacing_w_px,
                stirrup_combo_h_px,
                ls,
                store_key=sk,
            )
        if policy_x is not None:
            self._place_segment_stirrup_policy_combo(
                canvas,
                policy_x,
                y_top,
                hpx,
                int(slot),
                eid,
                policy_w_px,
                stirrup_combo_h_px,
                ls,
            )

    def _prepare_scheme_header_for_paint(self):
        u"""Canvas fijo encima del ScrollViewer; vaciar antes de repintar el esquema."""
        hc = getattr(self, "_scheme_header_canvas", None)
        if hc is None and self.window is not None:
            try:
                hc = self.window.FindName("TroceoSchemeHeaderCanvas")
                self._scheme_header_canvas = hc
            except Exception:
                hc = None
        if hc is not None:
            try:
                hc.Children.Clear()
            except Exception:
                pass
        return hc

    def _paint_fixed_scheme_headers(
        self,
        header_cv,
        scheme_cv_width,
        combo_left_x,
        spx_h,
        ls,
        stirrup_headers_active,
        has_stirrup_bar_combo,
        stirrup_diam_w_px,
        stirrup_spacing_w_px,
        stirrup_combo_inner_gap_px,
        nivel_value_left,
        shaft_x,
        shaft_w,
        bar_x0,
        long_combo_w=None,
        empalme_header_cx=None,
        troceo_empalme_header_cx=None,
        troceo_empalme_header_w_px=None,
        bubble_cx=None,
        lote_col_x=None,
        policy_x=None,
        policy_w_px=None,
        section_left_x=None,
        section_zone_w_px=None,
    ):
        u"""Encabezados centrados sobre controles; \u00d8 Long. sobre el combo (Tramo queda a la izquierda en el esquema)."""
        if header_cv is None:
            return
        hy = 0.5 * ls
        hz_t = 30
        try:
            header_cv.Width = float(scheme_cv_width)
        except Exception:
            pass

        def _center_tb(tb, cx, hy_f):
            if tb is None:
                return
            try:
                from System import Double
                from System.Windows import Size

                tb.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
                lw = float(tb.DesiredSize.Width)
                Canvas.SetLeft(tb, float(cx) - lw / 2.0)
            except Exception:
                pass
            try:
                Canvas.SetTop(tb, hy_f)
            except Exception:
                pass

        tb_esq = TextBlock()
        tb_esq.Text = u"Esquema"
        _apply_troceo_scheme_header_text_style(tb_esq, ls)
        try:
            from System import Double
            from System.Windows import Size

            tb_esq.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            lw_esq = float(tb_esq.DesiredSize.Width)
        except Exception:
            lw_esq = 48.0
        try:
            cx = float(shaft_x) + float(shaft_w) / 2.0
            Canvas.SetLeft(tb_esq, cx - lw_esq / 2.0)
        except Exception:
            Canvas.SetLeft(tb_esq, 0.0)
        Canvas.SetTop(tb_esq, hy)
        try:
            Panel.SetZIndex(tb_esq, hz_t)
        except Exception:
            pass
        header_cv.Children.Add(tb_esq)
        sdw = float(stirrup_diam_w_px or 0.0)
        ssw = float(stirrup_spacing_w_px or 0.0)
        hdr_nudge_x = max(2.0, 2.25 * float(ls))
        if stirrup_headers_active:
            if has_stirrup_bar_combo and sdw > 0.5:
                tbd = TextBlock()
                tbd.Text = u"\u00f8 Estribo"
                _apply_troceo_scheme_header_text_style(tbd, ls)
                try:
                    _center_tb(
                        tbd,
                        float(combo_left_x) + sdw / 2.0 + hdr_nudge_x,
                        hy,
                    )
                except Exception:
                    Canvas.SetLeft(tbd, float(combo_left_x))
                    Canvas.SetTop(tbd, hy)
                try:
                    Panel.SetZIndex(tbd, hz_t)
                except Exception:
                    pass
                header_cv.Children.Add(tbd)
            tbs = TextBlock()
            tbs.Text = u"Esp. Estribo"
            _apply_troceo_scheme_header_text_style(tbs, ls)
            try:
                if has_stirrup_bar_combo and sdw > 0.5 and ssw > 0.5:
                    _center_tb(
                        tbs,
                        float(spx_h) + ssw / 2.0 + hdr_nudge_x,
                        hy,
                    )
                elif ssw > 0.5:
                    _center_tb(
                        tbs,
                        float(combo_left_x) + ssw / 2.0 + hdr_nudge_x,
                        hy,
                    )
                else:
                    Canvas.SetLeft(tbs, float(spx_h))
                    Canvas.SetTop(tbs, hy)
            except Exception:
                Canvas.SetLeft(tbs, float(spx_h))
                Canvas.SetTop(tbs, hy)
            try:
                Panel.SetZIndex(tbs, hz_t)
            except Exception:
                pass
            header_cv.Children.Add(tbs)
            if lote_col_x is not None:
                try:
                    lcx = float(lote_col_x)
                    lote_w = _stirrup_lote_label_col_width_px(ls)
                    tbl = TextBlock()
                    tbl.Text = u"Lote"
                    _apply_troceo_scheme_header_text_style(tbl, ls)
                    _center_tb(tbl, lcx + lote_w / 2.0, hy)
                    try:
                        Panel.SetZIndex(tbl, hz_t)
                    except Exception:
                        pass
                    header_cv.Children.Add(tbl)
                except Exception:
                    pass
            if policy_x is not None and policy_w_px is not None:
                try:
                    pw = float(policy_w_px)
                    if pw > 0.5:
                        tbp = TextBlock()
                        tbp.Text = u"Pol\u00edtica"
                        _apply_troceo_scheme_header_text_style(tbp, ls)
                        _center_tb(tbp, float(policy_x) + pw / 2.0, hy)
                        try:
                            Panel.SetZIndex(tbp, hz_t)
                        except Exception:
                            pass
                        header_cv.Children.Add(tbp)
                except Exception:
                    pass
            if section_left_x is not None and section_zone_w_px is not None:
                try:
                    tbsct = TextBlock()
                    tbsct.Text = u"Secci\u00f3n"
                    _apply_troceo_scheme_header_text_style(tbsct, ls)
                    _center_tb(
                        tbsct,
                        float(section_left_x) + float(section_zone_w_px) / 2.0,
                        hy,
                    )
                    try:
                        Panel.SetZIndex(tbsct, hz_t)
                    except Exception:
                        pass
                    header_cv.Children.Add(tbsct)
                except Exception:
                    pass
        if troceo_empalme_header_cx is not None:
            try:
                tcx = float(troceo_empalme_header_cx)
            except Exception:
                tcx = None
            if tcx is not None:
                tb_tr = TextBlock()
                tb_tr.Text = u"Troceo"
                _apply_troceo_scheme_header_text_style(tb_tr, ls)
                try:
                    from System.Windows.Controls import ToolTip

                    tip_tr = ToolTip()
                    tip_tr.Content = (
                        u"Base: inicio de eje. Mitad altura: mitad del sólido; "
                        u"B con L(Ø); traslapo 50/50 en la junta."
                    )
                    tb_tr.ToolTip = tip_tr
                except Exception:
                    pass
                _center_tb(tb_tr, tcx, hy)
                try:
                    Panel.SetZIndex(tb_tr, hz_t)
                except Exception:
                    pass
                header_cv.Children.Add(tb_tr)
        if empalme_header_cx is not None:
            try:
                ecx = float(empalme_header_cx)
            except Exception:
                ecx = None
            if ecx is not None:
                tb_emp = TextBlock()
                tb_emp.Text = u"Empalme"
                _apply_troceo_scheme_header_text_style(tb_emp, ls)
                try:
                    from System.Windows.Controls import ToolTip

                    tip = ToolTip()
                    tip.Content = u"Referencia de troceo: casilla «Define Empalme»."
                    tb_emp.ToolTip = tip
                except Exception:
                    pass
                _center_tb(tb_emp, ecx, hy)
                try:
                    Panel.SetZIndex(tb_emp, hz_t)
                except Exception:
                    pass
                header_cv.Children.Add(tb_emp)
        tb_long = TextBlock()
        tb_long.Text = u"\u00f8 Long."
        _apply_troceo_scheme_header_text_style(tb_long, ls)
        try:
            from System import Double
            from System.Windows import Size

            tb_long.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            lw_long = float(tb_long.DesiredSize.Width)
            lcw = float(long_combo_w) if long_combo_w is not None else 0.0
            if lcw > 0.5:
                Canvas.SetLeft(
                    tb_long,
                    float(bar_x0) + lcw / 2.0 - lw_long / 2.0,
                )
            else:
                Canvas.SetLeft(tb_long, float(bar_x0))
        except Exception:
            Canvas.SetLeft(tb_long, float(bar_x0))
        Canvas.SetTop(tb_long, hy)
        try:
            Panel.SetZIndex(tb_long, hz_t)
        except Exception:
            pass
        header_cv.Children.Add(tb_long)
        tb_lvl = TextBlock()
        tb_lvl.Text = u"Nivel"
        _apply_troceo_scheme_header_text_style(tb_lvl, ls)
        try:
            from System import Double
            from System.Windows import Size

            tb_lvl.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            lw_lvl = float(tb_lvl.DesiredSize.Width)
        except Exception:
            lw_lvl = 36.0
        try:
            if bubble_cx is not None:
                Canvas.SetLeft(tb_lvl, float(bubble_cx) - lw_lvl / 2.0)
            else:
                Canvas.SetLeft(tb_lvl, float(nivel_value_left))
        except Exception:
            Canvas.SetLeft(tb_lvl, float(nivel_value_left))
        Canvas.SetTop(tb_lvl, hy)
        try:
            Panel.SetZIndex(tb_lvl, 30)
        except Exception:
            pass
        header_cv.Children.Add(tb_lvl)

    def _populate_blocks(self, skip_reveal=False):
        u"""Esquema vertical: fuste, estribos a la izquierda; casilla troceo junto al fuste y \u00d8 longitudinal a su derecha."""
        self._stop_blocks_reveal_animation()
        host = self._blocks_host
        if host is None:
            return
        host.Children.Clear()
        self._prepare_scheme_header_for_paint()
        self._borders_by_slot = {}
        self._define_empalme_checkboxes = {}
        models = getattr(self, u"_seg_models", None) or []
        n = len(models)
        wrap = Border()
        wrap.HorizontalAlignment = HorizontalAlignment.Center
        wrap.Margin = Thickness(0.0, 4.0, 0.0, 6.0)
        if n == 0:
            tb = TextBlock()
            tb.Text = u"No hay columnas para mostrar en el esquema."
            tb.Foreground = _BR_MUTED
            tb.TextWrapping = TextWrapping.Wrap
            wrap.Child = tb
            host.Children.Add(wrap)
            self._scheme_canvas = None
            self._bar_ladder_layout = None
            self._troceo_layout_scale = 1.0
            self._sync_diameter_panel()
            return
        ls = float(_TROCEO_LAYOUT_SCALE)
        self._troceo_layout_scale = ls
        found_h = 14.0 * ls
        top_pad = 6.0 * ls
        shaft_w = 20.0 * ls
        bar_ladder_w = 18.0 * ls
        foundation_w = 34.0 * ls
        shaft_x_base = float(_TROCEO_SHAFT_X_BASE_PX) * ls
        seg_base = float(_TROCEO_SEG_UNIFORM_BASE_PX) * ls
        max_stack = float(_MAX_TROCEO_STACK_PX) * ls
        vcap = getattr(self, "_viewport_stack_budget_px", None)
        if vcap is not None and float(vcap) > 40.0:
            max_stack = min(max_stack, float(vcap))
        min_seg = float(_TROCEO_SEG_UNIFORM_MIN_PX) * ls
        seg_uniform = seg_base
        if n < _MIN_TROCEO_SEGMENTS_FOR_SCALE:
            seg_uniform = max(
                seg_base,
                min(
                    max_stack / float(max(n, 1)),
                    max_stack / float(_MIN_TROCEO_SEGMENTS_FOR_SCALE),
                ),
            )
        elif n * seg_uniform > max_stack:
            seg_uniform = max(min_seg, max_stack / float(n))
        try:
            _pol_store = getattr(self, u"_col_stirrup_policy_store", None) or {}
            _need_l3_h = _stirrup_combo_compact_height_px(ls) * 3.0 + 10.0 * ls
            for _m in models:
                if (
                    _stirrup_policy_for_slot(_pol_store, _m.get(u"slot", 0))
                    == STIRRUP_POLICY_THIRDS_L3
                ):
                    seg_uniform = max(float(seg_uniform), float(_need_l3_h))
                    break
        except Exception:
            pass
        seg_px = [float(seg_uniform)] * n
        stack_h = float(n) * seg_uniform
        if stack_h > max_stack + 1e-6:
            seg_uniform = max(min_seg, max_stack / float(max(n, 1)))
            seg_px = [float(seg_uniform)] * n
            stack_h = float(n) * seg_uniform
        # Borde inferior del fuste = techo de la zapata (sin hueco).
        fy = top_pad + stack_h
        dims_cache = {}
        for m in models:
            eid = int(m[u"eid"])
            if eid not in dims_cache:
                dims_cache[eid] = _column_plan_dims_short_long_mm(m.get(u"elem"))
        max_rot_w, _max_rot_h = _measure_section_label_bounds(dims_cache, ls)
        tick_x0_b = shaft_x_base - 8.0 * ls
        bubble_r = 8.0 * ls
        # Menos hueco escalerilla-burbuja: acota el ancho útil y aleja el símbolo del borde derecho (menos scroll horizontal).
        _bubble_gap_after_ladder = 1.25 * ls
        # Fuste → «Define Empalme» → [separador] → etiqueta «Tramo n» → Ø longitudinal → franja burbuja
        long_cb_w = float(_stirrup_combo_row_width_px(ls, True))
        fuste_to_empalme_gap = 3.0 * ls
        between_empalme_long = max(4.0, 3.75 * ls)
        troceo_pol_w_px = _troceo_empalme_policy_combo_width_px(ls)
        troceo_pol_gap_px = max(4.0, 3.5 * ls)
        empalme_chk_reserve = max(108.0, 104.0 * ls)
        trom_label_col_w = max(40.0, 38.0 * ls)
        shaft_right_b = shaft_x_base + shaft_w
        troceo_pol_x_b = shaft_right_b + fuste_to_empalme_gap
        empalme_x_b = troceo_pol_x_b + troceo_pol_w_px + troceo_pol_gap_px
        after_checkbox_b = empalme_x_b + empalme_chk_reserve
        trom_label_x0_b = after_checkbox_b + between_empalme_long
        long_x_b = trom_label_x0_b + trom_label_col_w
        bar_x0_b = long_x_b
        bubble_cx_b = long_x_b + long_cb_w + bar_ladder_w + _bubble_gap_after_ladder + bubble_r
        fx_b = shaft_x_base + (shaft_w - foundation_w) / 2.0
        pad_dim = 6.0 * ls
        _has_stirrup_bar_combo = bool(
            self._col_stirrup_bar_store is not None and self._col_stirrup_bar_choices
        )
        if self._col_spacing_store is not None:
            stirrup_combo_w_px = _stirrup_combo_row_width_px(ls, _has_stirrup_bar_combo)
            stirrup_spacing_w_px = stirrup_combo_w_px
            stirrup_diam_w_px = (
                stirrup_combo_w_px if _has_stirrup_bar_combo else 0.0
            )
            stirrup_combo_h_px = _stirrup_combo_compact_height_px(ls)
            stirrup_combo_inner_gap_px = (
                float(_TROCEO_STIRRUP_COMBO_PAIR_GAP_PX) * ls
                if _has_stirrup_bar_combo
                else 0.0
            )
            combo_w_px = (
                stirrup_diam_w_px + stirrup_combo_inner_gap_px + stirrup_spacing_w_px
            )
            combo_gap_px = 10.0 * ls
            combo_left_reserve_px = 12.0 * ls
            stirrup_policy_w_px = _stirrup_policy_combo_width_px(ls)
            stirrup_lote_col_w_px = _stirrup_lote_label_col_width_px(ls)
            stirrup_policy_gap_px = max(
                _TROCEO_STIRRUP_INTER_COL_GAP_PX * ls, 6.0
            )
            stirrup_lote_gap_px = max(5.0, 4.5 * ls)
            _strip_pre = _troceo_stirrup_strip_x_layout(
                shaft_x_base,
                dims_cache,
                ls,
                combo_w_px,
            )
            stirrup_controls_w_px = (
                float(shaft_x_base) - float(_strip_pre[u"strip_left"])
            )
        else:
            stirrup_spacing_w_px = 0.0
            stirrup_diam_w_px = 0.0
            stirrup_combo_h_px = 0.0
            stirrup_combo_inner_gap_px = 0.0
            combo_w_px = 0.0
            combo_gap_px = 0.0
            combo_left_reserve_px = 0.0
            stirrup_policy_w_px = 0.0
            stirrup_lote_col_w_px = 0.0
            stirrup_policy_gap_px = 0.0
            stirrup_lote_gap_px = 0.0
            stirrup_controls_w_px = 0.0
            _strip_pre = None
        if self._col_spacing_store is not None and combo_w_px > 0.5:
            left_dim_b = (
                float(_strip_pre[u"strip_left"])
                - combo_gap_px
                - combo_left_reserve_px
            )
        else:
            section_label_clearance_px = (
                float(max_rot_w) + float(pad_dim) + max(6.0, 5.0 * ls)
            )
            left_dim_b = (
                shaft_x_base
                - section_label_clearance_px
                - combo_w_px
                - combo_gap_px
                - combo_left_reserve_px
            )
        content_left = min(tick_x0_b, left_dim_b, fx_b)
        content_right = max(
            bubble_cx_b + bubble_r + 3.0 * ls,
            fx_b + foundation_w,
            shaft_x_base + shaft_w,
        )
        content_width = max(content_right - content_left, 1.0)
        pad_side = 12.0 * ls
        cv = Canvas()
        cv.Width = max(content_width + 2.0 * pad_side, 220.0 * ls)
        offset_x = (cv.Width - content_width) / 2.0 - content_left
        cv.Height = fy + found_h + 10.0 * ls
        wrap.Child = cv
        host.Children.Add(wrap)
        shaft_x = shaft_x_base + offset_x
        bar_x0 = bar_x0_b + offset_x
        trom_label_x0 = trom_label_x0_b + offset_x
        bubble_cx = bubble_cx_b + offset_x
        fx = fx_b + offset_x
        section_left_x = None
        section_zone_w_px = None
        if self._col_spacing_store is not None and combo_w_px > 0.5:
            _strip = _troceo_stirrup_strip_x_layout(
                float(shaft_x),
                dims_cache,
                ls,
                combo_w_px,
            )
            policy_x = float(_strip[u"policy_x"])
            combo_left_x = float(_strip[u"combo_left_x"])
            lote_col_x = float(_strip[u"lote_col_x"])
            section_left_x = float(_strip[u"section_left"])
            section_zone_w_px = float(_strip[u"section_w"])
            stirrup_policy_w_px = float(_strip[u"policy_w"])
        else:
            policy_x = None
            lote_col_x = None
            combo_left_x = (
                shaft_x_base
                - max_rot_w
                - pad_dim
                - combo_gap_px
                - combo_w_px
                - combo_left_reserve_px
                + offset_x
            )
        self._scheme_section_label_left = section_left_x
        self._scheme_section_zone_w = section_zone_w_px
        self._scheme_canvas = cv
        spx_h = float(combo_left_x)
        if self._col_spacing_store is not None and combo_w_px > 0.5:
            if _has_stirrup_bar_combo and stirrup_diam_w_px > 0.5:
                spx_h = (
                    float(combo_left_x)
                    + float(stirrup_diam_w_px)
                    + float(stirrup_combo_inner_gap_px)
                )
        level_strings = []
        for slot in range(n - 1, -1, -1):
            m = models[slot]
            z_mm = float(m[u"z_mm"])
            z_m = z_mm / 1000.0
            level_strings.append(u"{0:.3f}".format(z_m))
        nivel_value_left = _level_numeric_labels_left_px(
            shaft_x,
            shaft_w,
            bubble_cx,
            bubble_r,
            ls,
            level_strings,
        )
        troceo_pol_x = troceo_pol_x_b + offset_x
        empalme_x = empalme_x_b + offset_x
        troceo_pol_header_cx = float(troceo_pol_x) + float(troceo_pol_w_px) / 2.0
        empalme_header_cx = (
            float(empalme_x) + float(empalme_chk_reserve) / 2.0
        )
        self._paint_fixed_scheme_headers(
            getattr(self, "_scheme_header_canvas", None),
            cv.Width,
            combo_left_x,
            spx_h,
            ls,
            bool(self._col_spacing_store is not None and combo_w_px > 0.5),
            _has_stirrup_bar_combo,
            stirrup_diam_w_px if self._col_spacing_store is not None else None,
            stirrup_spacing_w_px if self._col_spacing_store is not None else 0.0,
            stirrup_combo_inner_gap_px,
            nivel_value_left,
            shaft_x,
            shaft_w,
            bar_x0,
            long_cb_w,
            empalme_header_cx=empalme_header_cx,
            troceo_empalme_header_cx=troceo_pol_header_cx,
            troceo_empalme_header_w_px=troceo_pol_w_px,
            bubble_cx=bubble_cx,
            lote_col_x=lote_col_x,
            policy_x=policy_x,
            policy_w_px=(
                stirrup_policy_w_px if self._col_spacing_store is not None else None
            ),
            section_left_x=section_left_x,
            section_zone_w_px=section_zone_w_px,
        )
        _draw_foundation(cv, fx, fy, foundation_w, found_h)
        brd_list = []
        y = top_pad
        seg_layout = []
        for slot in range(n - 1, -1, -1):
            hpx = seg_px[slot]
            m = models[slot]
            eid = int(m[u"eid"])
            ss_dim, sl_dim = dims_cache[eid]
            z_mm = float(m[u"z_mm"])
            z_m = z_mm / 1000.0
            display_lev = u"{0:.3f}".format(z_m)
            _, bd_br = _shaft_brushes_for_tramo(1)
            brd = Border()
            brd.Tag = int(slot)
            brd.Width = shaft_w
            brd.Height = hpx
            try:
                from System.Windows.Media import Brushes

                brd.Background = Brushes.Transparent
            except Exception:
                brd.Background = _BR_ENTRY_BG
            brd.BorderBrush = bd_br
            _th0, _cr0 = _shaft_segment_border_and_radius(int(slot), n, ls, False)
            brd.BorderThickness = _th0
            brd.CornerRadius = _cr0
            try:
                brd.Opacity = 0.0
            except Exception:
                pass
            try:
                from System.Windows.Controls import ToolTip

                tip = ToolTip()
                h_real = float(m[u"height_mm"])
                if ss_dim is not None and sl_dim is not None:
                    tip.Content = (
                        u"Z = {0:.3f} m\nId {1}\nAlto tramo ≈ {2:.0f} mm\n{3}x{4}\n"
                        u"Referencia troceo: casilla «Define Empalme» junto al fuste."
                    ).format(z_m, int(eid), h_real, int(ss_dim), int(sl_dim))
                else:
                    tip.Content = (
                        u"Z = {0:.3f} m\nId {1}\nAlto tramo ≈ {2:.0f} mm\n"
                        u"Referencia troceo: casilla «Define Empalme» junto al fuste."
                    ).format(z_m, int(eid), h_real)
                brd.ToolTip = tip
            except Exception:
                pass
            Canvas.SetLeft(brd, shaft_x)
            Canvas.SetTop(brd, y)
            try:
                Panel.SetZIndex(brd, 12)
            except Exception:
                pass
            cv.Children.Add(brd)
            self._borders_by_slot[int(slot)] = brd
            brd_list.append(brd)
            _add_left_segment_dims_vertical(
                cv,
                shaft_x,
                y,
                hpx,
                ss_dim,
                sl_dim,
                3,
                layout_scale=ls,
                section_label_left_x=getattr(
                    self, u"_scheme_section_label_left", None
                ),
                section_zone_w=getattr(self, u"_scheme_section_zone_w", None),
            )
            self._place_troceo_empalme_policy_combo(
                cv,
                troceo_pol_x,
                y,
                hpx,
                int(slot),
                troceo_pol_w_px,
                stirrup_combo_h_px if self._col_spacing_store else _stirrup_combo_compact_height_px(ls),
                ls,
            )
            self._place_segment_define_empalme_checkbox(
                cv,
                empalme_x,
                max(52.0, float(empalme_chk_reserve) - 2.0 * ls),
                y,
                hpx,
                int(slot),
                ls,
            )
            if self._col_spacing_store is not None and combo_w_px > 0.5:
                self._place_column_stirrup_row_controls(
                    cv,
                    lote_col_x,
                    combo_left_x,
                    policy_x,
                    y,
                    hpx,
                    int(slot),
                    eid,
                    m.get(u"elem"),
                    stirrup_diam_w_px,
                    stirrup_spacing_w_px,
                    stirrup_combo_h_px,
                    stirrup_policy_w_px,
                    ls,
                    _has_stirrup_bar_combo,
                    stirrup_combo_inner_gap_px,
                    shaft_x,
                    shaft_w,
                )
            seg_layout.append((y, hpx, slot, display_lev))
            y += hpx
        slot_to_y_bottom = {}
        slot_to_y_span = {}
        for y_t, hp, sl, _ in seg_layout:
            sl_i = int(sl)
            y_top = float(y_t)
            y_bot = y_top + float(hp)
            slot_to_y_bottom[sl_i] = y_bot
            slot_to_y_span[sl_i] = (y_top, y_bot)
        self._adjust_define_empalme_checkboxes_vs_long_combo(
            seg_layout, slot_to_y_bottom, top_pad, stack_h, ls
        )
        try:
            sep_x = (
                float(shaft_x)
                + float(shaft_w)
                + fuste_to_empalme_gap
                + empalme_chk_reserve
                + 0.5 * between_empalme_long
            )
            _add_troceo_ref_long_vertical_separator(
                cv, sep_x, float(top_pad), float(top_pad) + float(stack_h), ls
            )
        except Exception:
            pass
        self._bar_ladder_layout = {
            u"bar_x0": float(bar_x0),
            u"top_pad": float(top_pad),
            u"stack_h": float(stack_h),
            u"bar_ladder_w": float(bar_ladder_w),
            u"shaft_x": float(shaft_x),
            u"shaft_w": float(shaft_w),
            u"slot_to_y_bottom": slot_to_y_bottom,
            u"slot_to_y_span": slot_to_y_span,
            u"layout_scale": float(ls),
            u"tramo_label_x0": float(trom_label_x0),
            u"tramo_label_col_w": float(trom_label_col_w),
            u"long_combo_w": float(long_cb_w),
        }
        try:
            self._refresh_all_shaft_styles()
        except Exception:
            pass
        zi = 20
        for y_top, hpx, slot, lev in seg_layout:
            # Cota Z en borde inferior del tramo (Y hacia abajo). Marcador tipo Revit: valor en m (3 dec.) encima de la línea.
            y_base = y_top + hpx
            head_z = zi
            _add_revit_level_head(
                cv,
                y_base,
                bubble_cx,
                bubble_r,
                0.0,
                float(cv.Width),
                head_z,
            )
            lab = TextBlock()
            lab.Text = u"{0}".format(lev)
            lab.Foreground = _BR_LEVEL_TEXT
            lab.FontSize = 10.5 * ls
            lab.FontWeight = FontWeights.SemiBold
            try:
                from System import Double
                from System.Windows import Size

                lab.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
                lw = float(lab.DesiredSize.Width)
                lh = float(lab.DesiredSize.Height)
            except Exception:
                lw, lh = 40.0, 14.0
            col_right = float(shaft_x) + float(shaft_w)
            line_zone_left = col_right + max(1.5 * ls, 2.5)
            bubble_left = float(bubble_cx) - float(bubble_r)
            # Texto a la izquierda del disco: bajar cerca de la linea (y_base); hueco un poco mayor respecto al simbolo.
            num_h_gap = max(2.5 * ls, 3.5)
            lab_left = max(line_zone_left, bubble_left - num_h_gap - lw)
            Canvas.SetLeft(lab, lab_left)
            gap_above_line = max(1.0 * ls, 1.5)
            Canvas.SetTop(lab, y_base - lh - gap_above_line)
            try:
                Panel.SetZIndex(lab, head_z + 10)
            except Exception:
                pass
            zi += 20
            cv.Children.Add(lab)
        self._update_troceo_datum_reference_lines()
        # Repoblar Ø longitudinal: _populate_blocks recrea el canvas y destruye los combos previos.
        self._sync_diameter_panel()
        if skip_reveal:
            for brd in brd_list:
                try:
                    brd.Opacity = 1.0
                except Exception:
                    pass
            self._schedule_scroll_scheme_to_bottom()
        else:
            self._start_blocks_reveal_animation(brd_list)

    def _troceo_cut_y_by_slot(self, sel_slots=None, lay=None):
        u"""Mapa slot \u2192 Y de corte en pantalla (coherente con l\u00edneas rojas)."""
        if lay is None:
            lay = getattr(self, "_bar_ladder_layout", None)
        if lay is None:
            return {}
        if sel_slots is None:
            sel_slots = getattr(self, "_sel_slots", None) or set()
        bot_map = lay.get(u"slot_to_y_bottom") or {}
        span_map = lay.get(u"slot_to_y_span") or {}
        out = {}
        for s in sel_slots or []:
            y = self._troceo_datum_y_px_for_slot(s, bot_map, span_map)
            if y is not None:
                try:
                    out[int(s)] = float(y)
                except Exception:
                    pass
        return out

    def _bar_tramo_bands_for_layout(self, sel_slots=None, lay=None):
        u"""Bandas de tramo del fuste usando cortes de empalme (entre l\u00edneas rojas)."""
        if lay is None:
            lay = getattr(self, "_bar_ladder_layout", None)
        if lay is None:
            return []
        if sel_slots is None:
            sel_slots = getattr(self, "_sel_slots", None) or set()
        try:
            top_pad = float(lay[u"top_pad"])
            stack_h = float(lay[u"stack_h"])
            bot_map = lay.get(u"slot_to_y_bottom") or {}
        except Exception:
            return []
        return _bar_tramo_y_bands_from_cuts(
            sel_slots,
            bot_map,
            top_pad,
            stack_h,
            slot_to_cut_y=self._troceo_cut_y_by_slot(sel_slots, lay),
        )

    def _troceo_datum_y_px_for_slot(self, slot, bot_map, span_map):
        u"""Y de la l\u00ednea roja: base del tramo (Base) o mitad visual (Mitad altura)."""
        try:
            sl = int(slot)
        except Exception:
            return None
        if self._troceo_empalme_policy_for_slot(sl) == TROCEO_EMPALME_POLICY_MID_AXIS:
            span = (span_map or {}).get(sl)
            if span is not None:
                try:
                    y_top = float(span[0])
                    y_bot = float(span[1])
                    return 0.5 * (y_top + y_bot)
                except Exception:
                    pass
        yb = (bot_map or {}).get(sl)
        if yb is not None:
            try:
                return float(yb)
            except Exception:
                pass
        return None

    def _update_troceo_datum_reference_lines(self):
        u"""L\u00ednea roja de referencia de troceo por tramo con empalme (base o mitad del fuste)."""
        cv = getattr(self, "_scheme_canvas", None)
        lay = getattr(self, "_bar_ladder_layout", None)
        if cv is None or lay is None:
            return
        _remove_troceo_datum_reference_lines(cv)
        try:
            sel = getattr(self, "_sel_slots", None) or set()
            if not sel:
                return
            ls = float(lay.get(u"layout_scale", 1.0))
            shaft_x = float(lay[u"shaft_x"])
            shaft_w = float(lay[u"shaft_w"])
            cv_w = float(cv.Width)
            bot_map = lay.get(u"slot_to_y_bottom") or {}
            span_map = lay.get(u"slot_to_y_span") or {}
            _xref_lo = max(0.0, shaft_x - 6.0 * ls)
            _xref_hi = min(cv_w, shaft_x + shaft_w + 6.0 * ls)
            for sl in sorted(int(s) for s in sel):
                y_ref = self._troceo_datum_y_px_for_slot(sl, bot_map, span_map)
                if y_ref is not None:
                    _add_troceo_datum_reference_line(
                        cv, float(y_ref), _xref_lo, _xref_hi, ls,
                    )
        except Exception:
            pass

    def _refresh_all_shaft_styles(self):
        u"""Colorea el fuste por subtramos entre cortes de empalme (l\u00edneas rojas)."""
        lay = getattr(self, "_bar_ladder_layout", None)
        cv = getattr(self, "_scheme_canvas", None)
        if not lay:
            return
        span_map = lay.get(u"slot_to_y_span") or {}
        sel = getattr(self, "_sel_slots", None) or set()
        bands = self._bar_tramo_bands_for_layout(sel, lay)
        n_seg = len(self._seg_models or [])
        ls = float(getattr(self, "_troceo_layout_scale", 1.0))
        cut_ys = sorted(set(self._troceo_cut_y_by_slot(sel, lay).values()))
        try:
            shaft_x = float(lay[u"shaft_x"])
            shaft_w = float(lay[u"shaft_w"])
        except Exception:
            shaft_x = shaft_w = None
        if cv is not None and shaft_x is not None:
            _paint_troceo_shaft_tramo_fills(
                cv, shaft_x, shaft_w, span_map, bands, cut_ys
            )
        try:
            from System.Windows.Media import Brushes

            _bg_clear = Brushes.Transparent
        except Exception:
            _bg_clear = _BR_ENTRY_BG
        for slot, brd in (self._borders_by_slot or {}).items():
            if brd is None:
                continue
            tno = _slot_to_tramo_number_from_bands(int(slot), bands, span_map)
            _, bd = _shaft_brushes_for_tramo(tno)
            th, cr = _shaft_segment_border_and_radius(int(slot), n_seg, ls, False)
            try:
                brd.Background = _bg_clear
                brd.BorderBrush = bd
                brd.BorderThickness = th
                brd.CornerRadius = cr
            except Exception:
                pass

    def _sync_shaft_tramo_band_labels(self):
        u"""N\u00fameros 1\u2026N centrados en el fuste por banda de troceo (coherente con la escalerilla)."""
        cv = getattr(self, "_scheme_canvas", None)
        lay = getattr(self, "_bar_ladder_layout", None)
        if cv is None or lay is None:
            return
        try:
            slot_bot = lay.get(u"slot_to_y_bottom") or {}
            top_pad = float(lay[u"top_pad"])
            stack_h = float(lay[u"stack_h"])
            shaft_x = float(lay[u"shaft_x"])
            shaft_w = float(lay[u"shaft_w"])
            ls = float(lay.get(u"layout_scale", 1.0))
            sel = getattr(self, "_sel_slots", None) or set()
            bands = self._bar_tramo_bands_for_layout(sel, lay)
        except Exception:
            return
        _remove_troceo_shaft_tramo_labels(cv)
        _draw_troceo_shaft_tramo_labels(cv, shaft_x, shaft_w, bands, ls)

    def _read_alternate_tramo_params(self):
        """Tramo inicial en base 1 (1 = base del pilar) y paso N (cada cuántos tramos se alterna)."""
        start_1 = 2
        step = 2
        tb_s = getattr(self, "_tb_alternate_start", None)
        tb_t = getattr(self, "_tb_alternate_step", None)
        if tb_s is not None:
            try:
                txt = tb_s.Text
                if txt is not None:
                    start_1 = int(float(unicode(txt).strip().replace(",", ".")))
            except Exception:
                pass
        if tb_t is not None:
            try:
                txt = tb_t.Text
                if txt is not None:
                    step = int(float(unicode(txt).strip().replace(",", ".")))
            except Exception:
                pass
        return max(1, int(start_1)), max(1, int(step))

    def _slots_alternate_pattern(self):
        """Índices ``slot`` (0 = tramo inferior) que siguen el patrón inicio/cada N."""
        n = len(self._seg_models or [])
        if n < 1:
            return set()
        start_tramo_1, step = self._read_alternate_tramo_params()
        if start_tramo_1 > n:
            return set()
        start_slot = start_tramo_1 - 1
        out = set()
        s = int(start_slot)
        while s < n:
            out.add(s)
            s += int(step)
        return out

    def _sync_define_empalme_checkboxes_from_sel_slots(self):
        """Alinea los checkboxes «Define Empalme» con ``_sel_slots`` (p. ej. tras Alternar referencia)."""
        try:
            self._troceo_empalme_cb_suppress = True
            for slot, chk in (self._define_empalme_checkboxes or {}).items():
                if chk is None:
                    continue
                try:
                    chk.IsChecked = int(slot) in self._sel_slots
                except Exception:
                    pass
        finally:
            self._troceo_empalme_cb_suppress = False

    def _on_define_empalme_checkbox_changed(self, sender, args):
        if getattr(self, "_troceo_empalme_cb_suppress", False):
            return
        self._finish_blocks_reveal_animation()
        try:
            chk = sender
            slot = int(chk.Tag)
        except Exception:
            return
        try:
            is_on = bool(chk.IsChecked)
        except Exception:
            is_on = False
        if is_on:
            self._sel_slots.add(slot)
        else:
            self._sel_slots.discard(slot)
        try:
            self._sync_troceo_empalme_policy_combo_state(int(slot))
        except Exception:
            pass
        if is_on:
            try:
                if self._sync_stirrup_policy_with_troceo_empalme(int(slot)):
                    self._sync_diameter_panel()
                    return
            except Exception:
                pass
        try:
            self._refresh_all_shaft_styles()
        except Exception:
            pass
        self._update_troceo_datum_reference_lines()
        try:
            self._sync_shaft_tramo_band_labels()
            self._refresh_bar_tramos_ladder()
        except Exception:
            pass
        self._sync_diameter_panel()

    def _on_alternate_sel(self, sender, args):
        self._finish_blocks_reveal_animation()
        pattern = self._slots_alternate_pattern()
        if not pattern:
            if TaskDialog is not None:
                try:
                    TaskDialog.Show(
                        u"Arainco: Esquema de troceo",
                        u"No hay tramos que coincidan con «desde» / «cada». "
                        u"Ajuste los valores (tramo 1 = base) o añada tramos al esquema.",
                    )
                except Exception:
                    pass
            return
        for s in sorted(pattern):
            if s in self._sel_slots:
                self._sel_slots.discard(s)
            else:
                self._sel_slots.add(s)
        self._sync_define_empalme_checkboxes_from_sel_slots()
        try:
            self._sync_all_troceo_empalme_policy_combos()
        except Exception:
            pass
        repainted = False
        for s in sorted(self._sel_slots):
            try:
                if self._sync_stirrup_policy_with_troceo_empalme(int(s)):
                    repainted = True
            except Exception:
                pass
        if repainted:
            self._sync_diameter_panel()
            return
        try:
            self._refresh_all_shaft_styles()
        except Exception:
            pass
        self._update_troceo_datum_reference_lines()
        try:
            self._sync_shaft_tramo_band_labels()
            self._refresh_bar_tramos_ladder()
        except Exception:
            pass
        self._sync_diameter_panel()

    def _wire_buttons(self):
        if self._embedded:
            b_ok = self._embed_btn_confirm
            b_cancel = self._embed_btn_cancel
            b_alt = self._embed_btn_alternate_sel
        else:
            b_ok = self.window.FindName("BtnConfirm")
            b_cancel = self.window.FindName("BtnCancel")
            b_alt = self.window.FindName("BtnAlternateSel")
            self.window.Closed += self._on_closed
        if b_ok is not None:
            b_ok.Click += self._on_confirm
        if b_cancel is not None:
            b_cancel.Click += self._on_cancel
        if b_alt is not None:
            b_alt.Click += self._on_alternate_sel
        self._wire_alt_numeric_steppers()

    def _wire_alt_numeric_steppers(self):
        try:
            w = self.window
            if w is None:
                return
            wire_wpf_numeric_stepper(
                w.FindName("TbTroceoAltStart"),
                w.FindName("BtnTroceoAltStartInc"),
                w.FindName("BtnTroceoAltStartDec"),
            )
            wire_wpf_numeric_stepper(
                w.FindName("TbTroceoAltStep"),
                w.FindName("BtnTroceoAltStepInc"),
                w.FindName("BtnTroceoAltStepDec"),
            )
        except Exception:
            pass

    def _on_closed(self, sender, args):
        if self._embedded:
            return
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

    def build_outcome_after_confirm(self):
        """Tras validar en ``_on_confirm`` (modo incrustado o no)."""
        return TroceoSchemeOutcome(
            cancelled=False,
            skip_no_cut=False,
            columns=self._selected_columns_ordered(),
            segment_rebar_bar_type_ids=list(self._pending_bar_type_ids)
            if self._pending_bar_type_ids
            else None,
            troceo_empalme_policy_by_column_id=(
                self._collect_troceo_empalme_policy_by_column_id()
            ),
        )

    def _on_cancel(self, sender, args):
        if self._embed_notify is not None:
            try:
                self._embed_notify("cancel")
            except Exception:
                pass
            return
        self.window.DialogResult = False
        try:
            self.window.Close()
        except Exception:
            pass

    def _on_confirm(self, sender, args):
        if not self._bar_choices:
            if TaskDialog is not None:
                try:
                    TaskDialog.Show(
                        u"Arainco: Esquema de troceo",
                        u"No hay RebarBarType en el proyecto.",
                    )
                except Exception:
                    pass
            return
        n_exp = len(self._sel_slots) + 1
        if len(self._diam_combos) != n_exp:
            if TaskDialog is not None:
                try:
                    TaskDialog.Show(
                        u"Arainco: Esquema de troceo",
                        u"Faltan tipos de barra por tramo (esperados {0}).".format(
                            n_exp
                        ),
                    )
                except Exception:
                    pass
            return
        ids = []
        for i, cb in enumerate(self._diam_combos):
            sel = cb.SelectedItem
            if sel is None:
                if TaskDialog is not None:
                    try:
                        TaskDialog.Show(
                            u"Arainco: Esquema de troceo",
                            u"Seleccione RebarBarType para el tramo {0}.".format(i + 1),
                        )
                    except Exception:
                        pass
                return
            try:
                tid = int(sel.Tag)
            except Exception:
                if TaskDialog is not None:
                    try:
                        TaskDialog.Show(
                            u"Arainco: Esquema de troceo",
                            u"No se pudo leer el tipo del tramo {0}.".format(i + 1),
                        )
                    except Exception:
                        pass
                return
            ids.append(tid)
        self._pending_bar_type_ids = ids
        if self._embed_notify is not None:
            try:
                self._embed_notify("confirm")
            except Exception:
                pass
            return
        self.window.DialogResult = True
        try:
            self.window.Close()
        except Exception:
            pass

    def _run_top_embed_collision_check(self, top_diam_mm, concrete_grade=None):
        u"""Chequeo de colisi\u00f3n del empotramiento superior de la pila.

        Devuelve ``True`` si el empotramiento choca con un s\u00f3lido estructural (se conserva),
        ``False`` si NO hay colisi\u00f3n (se revierte → sin extensi\u00f3n superior) y ``None`` si no
        fue posible realizar la prueba.
        """
        doc = getattr(self, u"_doc", None)
        models = getattr(self, u"_seg_models", None) or []
        if not doc or not models:
            return None
        cache = getattr(self, u"_cached_top_collision", None)
        if cache is None:
            self._cached_top_collision = {}
            cache = self._cached_top_collision
        diam_key = round(float(top_diam_mm), 1)
        if diam_key in cache:
            return cache[diam_key]
        result = None
        try:
            from column_reinforcement_layout_rps import (
                LAYOUT_EMBED_CONCRETE_GRADE,
                _resolved_traslape_embed_mm,
                _geometry_options_structure_solids,
                embed_stretch_collides_any_column_solids,
            )
            from Autodesk.Revit.DB import (
                FilteredElementCollector,
                BuiltInCategory,
                UnitUtils,
                UnitTypeId,
                XYZ,
            )
            grade_eff = concrete_grade or LAYOUT_EMBED_CONCRETE_GRADE
            embed_raw = _resolved_traslape_embed_mm(float(top_diam_mm), grade_eff) or 0.0
            embed_ft = UnitUtils.ConvertToInternalUnits(float(embed_raw), UnitTypeId.Millimeters)
            if embed_ft < 1e-9:
                cache[diam_key] = None
                return None
            z_base_mm = min(float(m[u"z_mm"]) for m in models)
            top_model = max(
                models,
                key=lambda m: float(m[u"z_mm"]) + float(m.get(u"height_mm") or 0.0),
            )
            z_top_mm = float(top_model[u"z_mm"]) + float(top_model.get(u"height_mm") or 0.0)
            fused_span_mm = z_top_mm - z_base_mm
            z_base_ft = UnitUtils.ConvertToInternalUnits(z_base_mm, UnitTypeId.Millimeters)
            fused_span_ft_val = UnitUtils.ConvertToInternalUnits(fused_span_mm, UnitTypeId.Millimeters)
            top_elem = top_model.get(u"elem")
            if top_elem is None:
                cache[diam_key] = None
                return None
            loc = top_elem.Location
            if hasattr(loc, u"Curve"):
                c = loc.Curve
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                cx = (float(p0.X) + float(p1.X)) / 2.0
                cy = (float(p0.Y) + float(p1.Y)) / 2.0
            elif hasattr(loc, u"Point"):
                cx = float(loc.Point.X)
                cy = float(loc.Point.Y)
            else:
                cache[diam_key] = None
                return None
            xyz_base = XYZ(cx, cy, z_base_ft)
            all_cols = list(
                FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType()
                .ToElements()
            )
            geom_opts = _geometry_options_structure_solids()
            contrib_ids = frozenset(int(m[u"eid"]) for m in models)
            result = bool(
                embed_stretch_collides_any_column_solids(
                    doc,
                    xyz_base,
                    fused_span_ft_val,
                    embed_ft,
                    float(top_diam_mm),
                    all_cols,
                    geom_opts,
                    contrib_ids,
                )
            )
        except Exception:
            result = None
        cache[diam_key] = result
        return result

    def _selected_columns_ordered(self):
        picked = self._sel_slots
        out = []
        for elem, z_mm, eid, slot in self._row_entries:
            if slot in picked:
                out.append(elem)
        return out

    def _attach_revit_owner_and_position(self):
        try:
            from System.Windows.Interop import WindowInteropHelper

            from revit_wpf_window_position import (
                position_wpf_window_top_left_at_active_view,
                revit_main_hwnd,
            )

            hwnd = revit_main_hwnd(self._uiapp)
            if hwnd:
                WindowInteropHelper(self.window).Owner = hwnd
            position_wpf_window_top_left_at_active_view(
                self.window,
                self._uidoc,
                hwnd,
            )
        except Exception:
            try:
                from System.Windows import WindowStartupLocation

                self.window.WindowStartupLocation = WindowStartupLocation.CenterScreen
            except Exception:
                pass

    def show_dialog(self):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, self.window)
        except Exception:
            pass
        self._attach_revit_owner_and_position()
        apply_troceo_scheme_window_maximized(self.window)
        try:
            self.window.Activate()
        except Exception:
            pass
        ok = self.window.ShowDialog()
        if not ok:
            return TroceoSchemeOutcome(cancelled=True)
        return TroceoSchemeOutcome(
            cancelled=False,
            skip_no_cut=False,
            columns=self._selected_columns_ordered(),
            segment_rebar_bar_type_ids=list(self._pending_bar_type_ids)
            if self._pending_bar_type_ids
            else None,
            troceo_empalme_policy_by_column_id=(
                self._collect_troceo_empalme_policy_by_column_id()
            ),
        )


def show_troceo_scheme_singleton(
    rows,
    uiapp=None,
    uidoc=None,
    doc=None,
    default_bar_diam_mm=12.0,
):
    """
    ``rows``: tuplas ``(elemento, z_mm, id [, height_mm [, level_name]])`` ordenadas
    por ``z_mm`` ascendente (o se reordenan en el controlador). ``height_mm`` y nombre
    de nivel son opcionales; si faltan, se infieren en la UI.
    ``doc``: ``Document`` Revit (si es None, se usa ``uidoc.Document``).
    """
    if doc is None and uidoc is not None:
        try:
            doc = uidoc.Document
        except Exception:
            doc = None
    try:
        existing = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
        if existing is not None:
            try:
                loaded = getattr(existing, "IsLoaded", False)
            except Exception:
                loaded = False
            if not loaded:
                try:
                    AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
                except Exception:
                    pass
                existing = None
        if existing is not None:
            try:
                existing.Activate()
                existing.Focus()
            except Exception:
                pass
            if TaskDialog is not None:
                try:
                    TaskDialog.Show(
                        u"Arainco: Armado Columnas",
                        u"La herramienta ya esta en ejecucion.",
                    )
                except Exception:
                    pass
            return None
    except Exception:
        pass
    if not rows:
        return TroceoSchemeOutcome(skip_no_cut=True, columns=[])
    ctrl = TroceoSchemeController(
        rows,
        uiapp=uiapp,
        uidoc=uidoc,
        doc=doc,
        default_bar_diam_mm=float(default_bar_diam_mm),
        parent_window=None,
        blocks_host=None,
        diam_host=None,
        btn_confirm=None,
        btn_cancel=None,
        embed_notify=None,
    )
    return ctrl.show_dialog()
