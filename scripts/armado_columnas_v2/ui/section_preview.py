# -*- coding: utf-8 -*-
"""Preview de sección de columna (planta): contorno, estribo perimetral, longitudinales."""

from __future__ import division

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Point, Size
from System.Windows.Controls import Canvas, TextBlock
from System.Windows.Media import (
    Color,
    PathFigure,
    LineSegment,
    ArcSegment,
    PathGeometry,
    PenLineJoin,
    RotateTransform,
    SolidColorBrush,
    SweepDirection,
    DoubleCollection,
)
from System.Windows.Shapes import Ellipse, Line, Path, Rectangle

from armado_columnas_v2.ui import layout as lay
from armado_columnas_v2.ui.wpf_brushes import brush_hex

PREVIEW_CANVAS_W = lay.SECTION_CTRL_WIDTH_PX
PREVIEW_CANVAS_H = lay.PREVIEW_CANVAS_H


def _canvas_dims(canvas):
    try:
        w = float(canvas.Width)
        cw = w if w > 1.0 else PREVIEW_CANVAS_W
    except Exception:
        cw = PREVIEW_CANVAS_W
    try:
        h = float(canvas.Height)
        ch = h if h > 1.0 else PREVIEW_CANVAS_H
    except Exception:
        ch = PREVIEW_CANVAS_H
    return cw, ch


def _fit_section_rect(w_cm, d_cm, canvas_w, canvas_h, pad_x=10.0, pad_top=4.0, label_h=12.0):
    max_w = canvas_w - pad_x * 2.0
    max_h = canvas_h - pad_top - label_h - 4.0
    b = max(1.0, float(w_cm))
    d = max(1.0, float(d_cm))
    if d >= b:
        sec_h = max_h
        sec_w = sec_h * (b / d)
        if sec_w > max_w:
            sec_w = max_w
            sec_h = sec_w * (d / b)
    else:
        sec_w = max_w
        sec_h = sec_w * (d / b)
        if sec_h > max_h:
            sec_h = max_h
            sec_w = sec_h * (b / d)
    ox = (canvas_w - sec_w) * 0.5
    oy = pad_top + (max_h - sec_h) * 0.5
    return ox, oy, sec_w, sec_h


def _bar_radius(diam_mm, scale):
    r = (float(diam_mm) / 18.0) * 1.45 * scale
    return max(1.2, min(3.2, r))


def _add_dot(canvas, x, y, r, fill_hex, stroke_hex, opacity=255):
    el = Ellipse()
    el.Width = r * 2.0
    el.Height = r * 2.0
    Canvas.SetLeft(el, x - r)
    Canvas.SetTop(el, y - r)
    el.Fill = brush_hex(fill_hex, opacity)
    el.Stroke = brush_hex(stroke_hex, min(255, opacity + 20))
    el.StrokeThickness = 0.8
    canvas.Children.Add(el)


def _add_outer_stirrup(canvas, x, y, w, h, stroke_brush, r_est):
    hook = min(3.2, w * 0.12, h * 0.12)
    fig = PathFigure()
    fig.StartPoint = Point(x + hook, y)
    fig.IsClosed = True
    segs = fig.Segments
    segs.Add(LineSegment(Point(x + w - hook, y), True))
    segs.Add(ArcSegment(
        Point(x + w, y + hook), Size(hook, hook), 0, False,
        SweepDirection.Clockwise, True,
    ))
    segs.Add(LineSegment(Point(x + w, y + h - hook), True))
    segs.Add(ArcSegment(
        Point(x + w - hook, y + h), Size(hook, hook), 0, False,
        SweepDirection.Clockwise, True,
    ))
    segs.Add(LineSegment(Point(x + hook, y + h), True))
    segs.Add(ArcSegment(
        Point(x, y + h - hook), Size(hook, hook), 0, False,
        SweepDirection.Clockwise, True,
    ))
    segs.Add(LineSegment(Point(x, y + hook), True))
    segs.Add(ArcSegment(
        Point(x + hook, y), Size(hook, hook), 0, False,
        SweepDirection.Clockwise, True,
    ))
    geom = PathGeometry()
    geom.Figures.Add(fig)
    path = Path()
    path.Data = geom
    path.Stroke = stroke_brush
    path.StrokeThickness = r_est
    path.StrokeLineJoin = PenLineJoin.Round
    path.Fill = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
    canvas.Children.Add(path)


def _add_dim_labels(canvas, ox, oy, sec_w, sec_h, w_cm, d_cm):
    dim_brush = brush_hex(u"#64748b")
    dim_size = 8.0

    h_lbl = TextBlock()
    h_lbl.Text = u"{0:.0f}".format(float(w_cm))
    h_lbl.FontSize = dim_size
    h_lbl.Foreground = dim_brush
    Canvas.SetLeft(h_lbl, ox + sec_w * 0.5 - 8.0)
    Canvas.SetTop(h_lbl, oy + sec_h + 6.0)
    canvas.Children.Add(h_lbl)

    v_lbl = TextBlock()
    v_lbl.Text = u"{0:.0f}".format(float(d_cm))
    v_lbl.FontSize = dim_size
    v_lbl.Foreground = dim_brush
    v_lbl.RenderTransformOrigin = Point(0.5, 0.5)
    v_lbl.RenderTransform = RotateTransform(-90.0)
    Canvas.SetLeft(v_lbl, ox - 18.0)
    Canvas.SetTop(v_lbl, oy + sec_h * 0.5 - 4.0)
    canvas.Children.Add(v_lbl)

    tick = brush_hex(u"#64748b", 140)
    for y_pos in (oy, oy + sec_h):
        ln = Line()
        ln.X1 = ox - 5.0
        ln.Y1 = y_pos
        ln.X2 = ox - 1.0
        ln.Y2 = y_pos
        ln.Stroke = tick
        ln.StrokeThickness = 0.7
        canvas.Children.Add(ln)
    v_axis = Line()
    v_axis.X1 = ox - 3.0
    v_axis.Y1 = oy
    v_axis.X2 = ox - 3.0
    v_axis.Y2 = oy + sec_h
    v_axis.Stroke = tick
    v_axis.StrokeThickness = 0.7
    canvas.Children.Add(v_axis)

    h_axis = Line()
    h_axis.X1 = ox
    h_axis.Y1 = oy + sec_h + 3.0
    h_axis.X2 = ox + sec_w
    h_axis.Y2 = oy + sec_h + 3.0
    h_axis.Stroke = tick
    h_axis.StrokeThickness = 0.7
    canvas.Children.Add(h_axis)


def _distribute_edge(count, a0, a1):
    if count <= 0:
        return []
    if count == 1:
        return [(a0 + a1) * 0.5]
    step = (a1 - a0) / float(count - 1)
    return [a0 + i * step for i in range(count)]


def _perimeter_bar_points(n_x, n_y, x0, y0, x1, y1):
    """Barras en perímetro rectangular (esquinas compartidas, una sola vez)."""
    n_x = max(2, int(n_x or 2))
    n_y = max(2, int(n_y or 2))
    pts = []
    xs = _distribute_edge(n_x, x0, x1)
    ys = _distribute_edge(n_y, y0, y1)

    # lados horizontales (incl. esquinas)
    for x in xs:
        pts.append((x, y0))
        pts.append((x, y1))
    # lados verticales sin esquina
    for y in ys[1:-1]:
        pts.append((x0, y))
        pts.append((x1, y))
    return pts


def section_meta_lines(member):
    if not member:
        return u"Sin elemento seleccionado."
    kind = member.get("kind") or u"column"
    w = float(member.get("widthCm") or 0)
    d = float(member.get("depthCm") or 0)
    lines = [
        u"{0} · {1}".format(
            member.get("label") or member.get("id"),
            member.get("typeName") or u"",
        ),
    ]
    elev_w = float(member.get("spanU_ft") or 0) * 30.48
    elev_h = float(member.get("spanV_ft") or 0) * 30.48

    if kind == u"foundation":
        lines.append(
            u"Fundación · planta {0:.0f}×{1:.0f} cm · alzado visto {2:.0f}×{3:.0f} cm".format(
                w, d, elev_w, elev_h
            )
        )
    elif kind == u"floor":
        lines.append(
            u"Losa · espesor {0:.0f} cm · alzado visto {1:.0f}×{2:.0f} cm".format(
                d, elev_w, elev_h
            )
        )
        base = member.get("levelBase")
        if base:
            lines.append(u"Nivel {0}".format(base))
    elif kind == u"beam":
        lines.append(
            u"Viga · sección {0:.0f}×{1:.0f} cm · alzado visto {2:.0f}×{3:.0f} cm".format(
                w, d, elev_w, elev_h
            )
        )
    else:
        nx = int(member.get("nBarsX") or 0)
        ny = int(member.get("nBarsY") or 0)
        dlong = int(member.get("diamLong") or 0)
        dest = int(member.get("diamEstribo") or 0)
        total = max(0, 2 * (nx + ny) - 4)
        lines.append(
            u"Sección {0:.0f}×{1:.0f} cm · long. {2}∅{3} · estribo ∅{4}".format(
                w, d, total, dlong, dest
            )
        )
        base = member.get("levelBase")
        top = member.get("levelTop")
        if base or top:
            lines.append(u"Alzado {0} → {1}".format(base or u"—", top or u"—"))
    return u"\n".join(lines)


def draw_section_preview(canvas, column):
    if canvas is None:
        return u""
    canvas.Children.Clear()
    if not column:
        return u""

    kind = column.get("kind") or u"column"
    is_column = kind == u"column"

    cw, ch = _canvas_dims(canvas)
    w_cm = float(column.get("widthCm") or 50.0)
    d_cm = float(column.get("depthCm") or 50.0)
    if kind == u"foundation":
        try:
            elev_w_cm = float(column.get("spanU_ft") or 0) * 30.48
            if d_cm < 1.0 or abs(d_cm - float(column.get("heightM") or 0) * 100.0) < 5.0:
                d_cm = max(w_cm, elev_w_cm)
            if w_cm < 1.0:
                w_cm = elev_w_cm or 80.0
        except Exception:
            pass
    elif kind == u"floor":
        # Preview: canto = espesor (depthCm), ancho simbólico
        if d_cm < 1.0:
            d_cm = float(column.get("spanV_ft") or 0.5) * 30.48
        if w_cm < 1.0:
            w_cm = max(d_cm * 4.0, 40.0)
    elif kind == u"beam":
        if w_cm < 1.0:
            w_cm = 30.0
        if d_cm < 1.0:
            d_cm = float(column.get("spanV_ft") or 0) * 30.48 or 50.0

    ox, oy, sec_w, sec_h = _fit_section_rect(w_cm, d_cm, cw, ch)

    cover_ratio = max(3.5, min(sec_w, sec_h) * 0.11)
    inner_x = ox + cover_ratio
    inner_y = oy + cover_ratio
    inner_w = sec_w - cover_ratio * 2.0
    inner_h = sec_h - cover_ratio * 2.0
    st_inset = 2.2
    st_x = inner_x + st_inset
    st_y = inner_y + st_inset
    st_w = inner_w - st_inset * 2.0
    st_h = inner_h - st_inset * 2.0
    scale = sec_h / max(d_cm, 1.0)

    stroke_by_kind = {
        u"column": u"#5bb8d4",
        u"foundation": u"#d4a574",
        u"beam": u"#7d9b8a",
        u"floor": u"#8b9cb3",
    }
    fill_by_kind = {
        u"column": u"#0a1620",
        u"foundation": u"#1a140c",
        u"beam": u"#0e1612",
        u"floor": u"#0e1218",
    }

    outer = Rectangle()
    outer.Width = sec_w
    outer.Height = sec_h
    Canvas.SetLeft(outer, ox)
    Canvas.SetTop(outer, oy)
    outer.RadiusX = 1.8
    outer.RadiusY = 1.8
    outer.Stroke = brush_hex(stroke_by_kind.get(kind, u"#5bb8d4"))
    outer.StrokeThickness = 1.3
    outer.Fill = brush_hex(fill_by_kind.get(kind, u"#0a1620"), 220)
    canvas.Children.Add(outer)

    inner = Rectangle()
    inner.Width = inner_w
    inner.Height = inner_h
    Canvas.SetLeft(inner, inner_x)
    Canvas.SetTop(inner, inner_y)
    inner.Stroke = brush_hex(u"#94a3b8", 56)
    inner.StrokeThickness = 0.6
    inner.StrokeDashArray = DoubleCollection([2.5, 2.0])
    inner.Fill = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
    canvas.Children.Add(inner)

    # Solo armado esquemático en columnas (rail aún en diseño para otros)
    if is_column:
        r_est = max(0.9, float(column.get("diamEstribo") or 10) / 9.0)
        _add_outer_stirrup(canvas, st_x, st_y, st_w, st_h, brush_hex(u"#34d399"), r_est)

        bar_pad = max(2.5, cover_ratio * 0.55)
        x0 = st_x + bar_pad
        x1 = st_x + st_w - bar_pad
        y0 = st_y + bar_pad
        y1 = st_y + st_h - bar_pad
        nx = int(column.get("nBarsX") or 4)
        ny = int(column.get("nBarsY") or 4)
        dlong = int(column.get("diamLong") or 22)
        r_bar = _bar_radius(dlong, scale)
        for px, py in _perimeter_bar_points(nx, ny, x0, y0, x1, y1):
            _add_dot(canvas, px, py, r_bar, u"#22d3ee", u"#0891b2", 255)

    _add_dim_labels(canvas, ox, oy, sec_w, sec_h, w_cm, d_cm)
    return section_meta_lines(column)
