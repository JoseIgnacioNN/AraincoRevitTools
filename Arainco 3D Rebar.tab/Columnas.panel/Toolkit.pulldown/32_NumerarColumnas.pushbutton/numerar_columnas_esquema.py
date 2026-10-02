# -*- coding: utf-8 -*-
"""
Esquemas simplificados de torres — paleta alineada con troceo_scheme (Armado Columnas portable).
"""

import clr

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Windows import (
    CornerRadius,
    FontStyles,
    FontWeights,
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    TextAlignment,
    TextTrimming,
    TextWrapping,
    Thickness,
    VerticalAlignment,
)
from System.Windows.Controls import (
    Border,
    Canvas,
    Grid,
    Orientation,
    RowDefinition,
    StackPanel,
    TextBlock,
    ToolTipService,
)
from System.Windows.Shapes import Rectangle
from System.Windows.Media import SolidColorBrush, Color

_CARD_WIDTH = 156.0
_CARD_CANVAS_H_DEFAULT = 112.0
_SHAFT_W = 44.0
_SHAFT_X = 16.0
_PAD = 6.0
_MIN_SEG_PX = 14.0
_CHROME_BASE_PX = 72.0
_LEGEND_LINE_PX = 15.0


def _hex_brush(hex_str):
    h = (hex_str or "#000000").lstrip("#")
    return SolidColorBrush(Color.FromRgb(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)))


_BR_FOUND = _hex_brush("#0c1824")
_BR_SIN_FUND = _hex_brush("#0f172a")
_BR_COL = tuple(
    _hex_brush(h)
    for h in (
        "#152535",
        "#1a3548",
        "#1d3f4a",
        "#243044",
        "#1a4038",
        "#2a3050",
    )
)
_BR_STROKE = _hex_brush("#475569")
_BR_TEXT = _hex_brush("#E8F4F8")
_BR_MUTED = _hex_brush("#94a3b8")
_BR_LOTE = _hex_brush("#E8F4F8")
_BR_BADGE = _hex_brush("#94a3b8")
_BR_CARD_BG = _hex_brush("#0f172a")
_BR_CARD_BD = _hex_brush("#334155")
_BR_CANVAS_BG = _hex_brush("#050a10")


def card_width_px():
    return _CARD_WIDTH


def _legend_row_count(segmentos):
    n = len(segmentos or [])
    max_rows = 5
    extra = max(0, n - max_rows)
    return min(n, max_rows) + (1 if extra > 0 else 0)


def _canvas_height_for_card(card_height_px, segmentos):
    if not card_height_px or card_height_px < 120:
        return _CARD_CANVAS_H_DEFAULT
    rows = _legend_row_count(segmentos)
    legend_h = rows * _LEGEND_LINE_PX + 6.0
    canvas_h = float(card_height_px) - _CHROME_BASE_PX - legend_h
    return max(100.0, canvas_h)


def _segment_badge(kind, tramo):
    if kind == u"fundacion":
        return u"F"
    if kind == u"sin_fundacion":
        return u"-"
    try:
        return u"{0}".format(int(tramo))
    except Exception:
        return u"C"


def pintar_torre_en_canvas(canvas, segmentos, width, height):
    """Torre solo con rectángulos (leyenda debajo)."""
    if canvas is None:
        return
    canvas.Children.Clear()
    try:
        canvas.Width = float(width)
        canvas.Height = float(height)
    except Exception:
        pass

    segs = list(segmentos or [])
    if not segs:
        return

    total_h = sum(max(0.05, float(s.get(u"h_ft") or 0.05)) for s in segs)
    if total_h < 0.01:
        total_h = 1.0

    usable_h = float(height) - 2.0 * _PAD
    y_bottom = float(height) - _PAD
    col_idx = 0

    for seg in segs:
        kind = seg.get(u"kind") or u"columna"
        h_ft = max(0.05, float(seg.get(u"h_ft") or 0.05))
        h_px = max(_MIN_SEG_PX, (h_ft / total_h) * usable_h)
        y_top = y_bottom - h_px

        rect = Rectangle()
        rect.Width = _SHAFT_W
        rect.Height = h_px
        if kind == u"fundacion":
            rect.Fill = _BR_FOUND
        elif kind == u"sin_fundacion":
            rect.Fill = _BR_SIN_FUND
        else:
            rect.Fill = _BR_COL[col_idx % len(_BR_COL)]
            col_idx += 1
        rect.Stroke = _BR_STROKE
        rect.StrokeThickness = 1.0
        rect.RadiusX = 2.0
        rect.RadiusY = 2.0
        Canvas.SetLeft(rect, _SHAFT_X)
        Canvas.SetTop(rect, y_top)
        canvas.Children.Add(rect)

        if h_px >= 18.0:
            badge = TextBlock()
            badge.Text = _segment_badge(kind, seg.get(u"tramo"))
            badge.FontSize = 9.0
            badge.FontWeight = FontWeights.SemiBold
            badge.Foreground = _BR_BADGE
            badge.TextAlignment = TextAlignment.Center
            badge.Width = _SHAFT_W
            Canvas.SetLeft(badge, _SHAFT_X)
            Canvas.SetTop(badge, y_top + max(2.0, (h_px - 12.0) / 2.0))
            canvas.Children.Add(badge)

        y_bottom = y_top


def _leyenda_segmentos(segmentos):
    panel = StackPanel()
    panel.Margin = Thickness(0, 8, 0, 0)
    segs = list(segmentos or [])
    max_rows = 5
    extra = max(0, len(segs) - max_rows)
    if extra > 0:
        segs = segs[:max_rows]
    for seg in segs:
        kind = seg.get(u"kind") or u"columna"
        short = seg.get(u"label") or u"?"
        full = seg.get(u"label_full") or short
        prefix = _segment_badge(kind, seg.get(u"tramo"))
        row = TextBlock()
        row.Text = u"{0}  {1}".format(prefix, short)
        row.FontSize = 10.0
        row.Foreground = _BR_MUTED
        row.TextTrimming = TextTrimming.CharacterEllipsis
        row.TextWrapping = TextWrapping.NoWrap
        row.MaxWidth = _CARD_WIDTH - 4.0
        row.Margin = Thickness(0, 1, 0, 1)
        row.ToolTip = full
        try:
            ToolTipService.SetShowDuration(row, 12000)
        except Exception:
            pass
        panel.Children.Add(row)
    if extra > 0:
        more = TextBlock()
        more.Text = u"+ {0} tramo(s) más (tooltip tarjeta)".format(extra)
        more.FontSize = 9.0
        more.Foreground = _BR_MUTED
        more.FontStyle = FontStyles.Italic
        panel.Children.Add(more)
    return panel


def _tooltip_tarjeta_completa(lote_no, torres_count, segmentos):
    lines = [u"Lote {0} — {1} torre(s)".format(int(lote_no), int(torres_count))]
    for seg in segmentos or []:
        kind = seg.get(u"kind")
        full = seg.get(u"label_full") or seg.get(u"label") or u"?"
        if kind == u"fundacion":
            lines.append(u"Fundación: {0}".format(full))
        elif kind == u"sin_fundacion":
            lines.append(u"Sin fundación")
        else:
            lines.append(u"Tramo {0}: {1}".format(seg.get(u"tramo") or u"?", full))
    return u"\n".join(lines)


def crear_tarjeta_lote(lote_no, torres_count, segmentos, card_height_px=None):
    segs = list(segmentos or [])
    card_w = _CARD_WIDTH + 20.0
    canvas_h = _canvas_height_for_card(card_height_px, segs)

    outer = Border()
    outer.Margin = Thickness(0, 0, 10, 0)
    outer.Padding = Thickness(10, 10, 10, 10)
    outer.Background = _BR_CARD_BG
    outer.BorderBrush = _BR_CARD_BD
    outer.BorderThickness = Thickness(1)
    outer.CornerRadius = CornerRadius(4.0)
    outer.Width = card_w
    outer.MinWidth = card_w
    outer.ToolTip = _tooltip_tarjeta_completa(lote_no, torres_count, segs)

    if card_height_px and card_height_px > 100:
        outer.Height = float(card_height_px)
        outer.MinHeight = float(card_height_px)
        outer.VerticalAlignment = VerticalAlignment.Stretch

    root = Grid()
    root.RowDefinitions.Add(RowDefinition())  # título
    root.RowDefinitions.Add(RowDefinition())  # esquema
    root.RowDefinitions.Add(RowDefinition())  # leyenda
    root.RowDefinitions.Add(RowDefinition())  # contador
    root.RowDefinitions[0].Height = GridLength(1, GridUnitType.Auto)
    root.RowDefinitions[1].Height = GridLength(canvas_h, GridUnitType.Pixel)
    root.RowDefinitions[2].Height = GridLength(1, GridUnitType.Auto)
    root.RowDefinitions[3].Height = GridLength(1, GridUnitType.Auto)

    tb_lote = TextBlock()
    tb_lote.Text = u"Lote {0}".format(int(lote_no))
    tb_lote.FontSize = 13.0
    tb_lote.FontWeight = FontWeights.SemiBold
    tb_lote.Foreground = _BR_LOTE
    tb_lote.HorizontalAlignment = HorizontalAlignment.Center
    tb_lote.Margin = Thickness(0, 0, 0, 8)
    Grid.SetRow(tb_lote, 0)
    root.Children.Add(tb_lote)

    cv = Canvas()
    cv.Width = _CARD_WIDTH
    cv.Height = canvas_h
    cv.MinHeight = canvas_h
    cv.Background = _BR_CANVAS_BG
    cv.ClipToBounds = True
    cv.HorizontalAlignment = HorizontalAlignment.Center
    cv.VerticalAlignment = VerticalAlignment.Center
    pintar_torre_en_canvas(cv, segs, _CARD_WIDTH, canvas_h)
    Grid.SetRow(cv, 1)
    root.Children.Add(cv)

    legend = _leyenda_segmentos(segs)
    Grid.SetRow(legend, 2)
    root.Children.Add(legend)

    tb_cnt = TextBlock()
    tb_cnt.Text = u"\u00d7 {0} torre{1}".format(
        int(torres_count),
        u"s" if int(torres_count) != 1 else u"",
    )
    tb_cnt.FontSize = 10.0
    tb_cnt.Foreground = _BR_MUTED
    tb_cnt.HorizontalAlignment = HorizontalAlignment.Center
    tb_cnt.Margin = Thickness(0, 8, 0, 0)
    Grid.SetRow(tb_cnt, 3)
    root.Children.Add(tb_cnt)

    outer.Child = root
    return outer


def poblar_galeria_horizontal(host_panel, lotes, card_height_px=None):
    if host_panel is None:
        return
    host_panel.Children.Clear()

    if card_height_px and card_height_px > 80:
        try:
            host_panel.MinHeight = float(card_height_px)
            host_panel.Height = float(card_height_px)
            host_panel.VerticalAlignment = VerticalAlignment.Stretch
        except Exception:
            pass

    if not lotes:
        tb = TextBlock()
        tb.Text = u"No hay lotes. Pulse \u00abAnalizar proyecto\u00bb."
        tb.Foreground = _BR_MUTED
        tb.FontSize = 11.0
        tb.Margin = Thickness(12, 24, 12, 12)
        host_panel.Children.Add(tb)
        return

    sorted_lotes = sorted(lotes, key=lambda x: int(x.get(u"lote") or 0))
    for item in sorted_lotes:
        try:
            card = crear_tarjeta_lote(
                item[u"lote"],
                item.get(u"torres_count") or 0,
                item.get(u"segmentos") or [],
                card_height_px=card_height_px,
            )
            if card_height_px and card_height_px > 80:
                try:
                    card.VerticalAlignment = VerticalAlignment.Stretch
                except Exception:
                    pass
            host_panel.Children.Add(card)
        except Exception as ex:
            err = TextBlock()
            err.Text = u"Lote {0}: {1}".format(item.get(u"lote"), ex)
            err.Foreground = _BR_MUTED
            err.FontSize = 10.0
            err.Margin = Thickness(8, 8, 8, 8)
            err.TextWrapping = TextWrapping.Wrap
            err.MaxWidth = _CARD_WIDTH
            host_panel.Children.Add(err)


def measure_gallery_card_height(scroll_viewer, fallback=320.0):
    """Altura útil para tarjetas según el ScrollViewer de la galería."""
    if scroll_viewer is None:
        return float(fallback)
    try:
        scroll_viewer.UpdateLayout()
    except Exception:
        pass
    h = 0.0
    try:
        h = float(scroll_viewer.ActualHeight)
    except Exception:
        pass
    if h < 80.0:
        try:
            h = float(scroll_viewer.RenderSize.Height)
        except Exception:
            pass
    if h < 80.0:
        h = float(fallback)
    return max(120.0, h - 4.0)
