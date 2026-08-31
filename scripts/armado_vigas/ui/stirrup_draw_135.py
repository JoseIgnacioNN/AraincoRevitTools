# -*- coding: utf-8 -*-
"""Estribo / traba en canvas 2D — estilo visual Armado Muros V3.

Port de ``Stirrup135Path`` + ``Hook135TrabaEnd`` de armado_muros_preview_ui:
esquinas redondeadas, ganchos 135° y margen exterior = radio barra + pad.
"""

import math

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Point as WpfPoint, Size as WpfSize
from System.Windows.Controls import Canvas, Panel
from System.Windows.Media import (
    ArcSegment,
    LineSegment,
    PathFigure,
    PathGeometry,
    PenLineCap,
    PenLineJoin,
    SolidColorBrush,
    SweepDirection,
)
from System.Windows.Shapes import Line, Path as WpfPath

from armado_vigas.ui.wpf_controls import brush_hex
from armado_vigas.ui.net_ui import freeze_freezable

# Constantes alineadas con Muros (preview cabezal)
BAR_DRAW_R_PX = 3.5
STIRRUP_MARGIN_PAD_PX = 4.0
STIRRUP_THICK = 2.2
TIE_THICK = 1.6
COLOR_ESTRIBO = u"#00B450"
COLOR_TRABA = u"#F59E0B"


def stirrup_margin_px(bar_r=None, pad=None):
    r = float(BAR_DRAW_R_PX if bar_r is None else bar_r)
    p = float(STIRRUP_MARGIN_PAD_PX if pad is None else pad)
    return r + p


def polar_px(cx, cy, deg, rad):
    a = float(deg) * math.pi / 180.0
    return (
        float(cx) + float(rad) * math.cos(a),
        float(cy) + float(rad) * math.sin(a),
    )


def add_stroke_path(canv, start, ops, brush, thick, z_index=21, round_caps=True, dash=None):
    """Path abierto: ops = ('L', x, y) | ('A', r, sweep_cw, x, y).

    ``dash``: secuencia de longitudes (p. ej. (5, 4)) para StrokeDashArray.
    """
    if canv is None or not ops:
        return
    try:
        from System.Windows.Media import DoubleCollection
    except Exception:
        DoubleCollection = None
    try:
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
                    SweepDirection.Clockwise if sweep_cw else SweepDirection.Counterclockwise
                )
                fig.Segments.Add(arc)
        geo = PathGeometry()
        geo.Figures.Add(fig)
        freeze_freezable(geo)
        path = WpfPath()
        path.Data = geo
        path.Stroke = brush
        path.StrokeThickness = float(thick)
        path.Fill = None
        if dash and DoubleCollection is not None:
            try:
                arr = DoubleCollection()
                for v in dash:
                    arr.Add(float(v))
                path.StrokeDashArray = arr
            except Exception:
                pass
        if round_caps:
            path.StrokeStartLineCap = PenLineCap.Round
            path.StrokeEndLineCap = PenLineCap.Round
            path.StrokeLineJoin = PenLineJoin.Round
        try:
            Panel.SetZIndex(path, int(z_index))
        except Exception:
            pass
        canv.Children.Add(path)
        return path
    except Exception:
        return None


def draw_stirrup_135(
    canv, left, top, right, bot, bar_cx, bar_cy, wrap_r,
    brush, thick, tip_left=True, z_index=21, dash=None,
):
    """Estribo cerrado con esquinas r + ganchos 135° (Stirrup135Path)."""
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
        return polar_px(bc_x, bc_y, deg, R if rad is None else rad)

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

    add_stroke_path(canv, north, body_ops, brush, thick, z_index, dash=dash)
    add_stroke_path(canv, north, hook_top_ops, brush, thick, z_index, dash=dash)
    add_stroke_path(canv, side_start, hook_side_ops, brush, thick, z_index, dash=dash)


def draw_hook_135_traba_end(
    canv, bar_cx, bar_cy, wrap_r, tip_left, end, brush, thick, z_index=22, dash=None,
):
    """Un gancho 135° en extremo de traba (Hook135TrabaEnd). end: top/bottom/left/right."""
    R = max(3.0, float(wrap_r))
    tail_len = max(18.0, float(thick) * 8.0)
    sqrt_half = math.sqrt(0.5)
    bc_x, bc_y = float(bar_cx), float(bar_cy)

    def _pol(deg):
        return polar_px(bc_x, bc_y, deg, R)

    end_key = unicode(end or u"top")
    if end_key == u"top":
        start = _pol(0) if tip_left else _pol(180)
        exit_pt = _pol(225) if tip_left else _pol(315)
        tx = (-sqrt_half) if tip_left else sqrt_half
        ty = sqrt_half
        sweep_cw = not tip_left
    elif end_key == u"bottom":
        start = _pol(0) if tip_left else _pol(180)
        exit_pt = _pol(135) if tip_left else _pol(45)
        tx = (-sqrt_half) if tip_left else sqrt_half
        ty = -sqrt_half
        sweep_cw = bool(tip_left)
    elif end_key == u"left":
        start = _pol(90)
        exit_pt = _pol(225)
        tx = sqrt_half
        ty = -sqrt_half
        sweep_cw = True
    else:
        start = _pol(90)
        exit_pt = _pol(315)
        tx = -sqrt_half
        ty = -sqrt_half
        sweep_cw = False

    ops = [
        (u"A", R, sweep_cw, exit_pt[0], exit_pt[1]),
        (u"L", exit_pt[0] + tx * tail_len, exit_pt[1] + ty * tail_len),
    ]
    add_stroke_path(canv, start, ops, brush, thick, z_index, dash=dash)


def draw_estribo_bbox(
    canv, points_xy, margin=None, brush=None, thick=None, tip_left=True, dashed=False,
):
    """Estribo 135° alrededor de puntos {(x,y)} o dicts con x/y."""
    if not points_xy or len(points_xy) < 2:
        return None
    pts = []
    for p in points_xy:
        if isinstance(p, dict):
            pts.append((float(p[u"x"]), float(p[u"y"])))
        else:
            pts.append((float(p[0]), float(p[1])))
    m = float(stirrup_margin_px() if margin is None else margin)
    cxs = [p[0] for p in pts]
    cys = [p[1] for p in pts]
    left = min(cxs) - m
    right = max(cxs) + m
    top = min(cys) - m
    bot = max(cys) + m
    if right - left < 4.0 or bot - top < 4.0:
        return None
    # Barra de esquina del gancho (TR si tip_left)
    max_x = max(cxs)
    min_y = min(cys)
    bar_cx, bar_cy = max_x, min_y
    best = 1e18
    for x, y in pts:
        score = (max_x - x) + (y - min_y)
        if score < best:
            best = score
            bar_cx, bar_cy = x, y
    if not tip_left:
        min_x = min(cxs)
        best = 1e18
        for x, y in pts:
            score = (x - min_x) + (y - min_y)
            if score < best:
                best = score
                bar_cx, bar_cy = x, y
    br = brush if brush is not None else brush_hex(COLOR_ESTRIBO)
    th = float(STIRRUP_THICK if thick is None else thick)
    dash = (5.0, 4.0) if dashed else None
    draw_stirrup_135(
        canv, left, top, right, bot, bar_cx, bar_cy, m, br, th, tip_left,
        z_index=21, dash=dash,
    )
    return (left, top, right - left, bot - top)


def draw_traba_vertical_column(
    canv, x_sup, y_sup, x_inf, y_inf,
    margin=None, brush=None, thick=None, tip_left=True, dashed=False,
):
    """Traba vertical viga: línea offset + ganchos 135° en barras SUP/INF."""
    m = float(stirrup_margin_px() if margin is None else margin)
    th = float(TIE_THICK if thick is None else thick)
    br = brush if brush is not None else brush_hex(COLOR_TRABA)
    cx = 0.5 * (float(x_sup) + float(x_inf))
    y_min = min(float(y_sup), float(y_inf))
    y_max = max(float(y_sup), float(y_inf))
    bx_top = float(x_sup)
    by_top = float(y_sup)
    bx_bot = float(x_inf)
    by_bot = float(y_inf)
    side = 1.0 if tip_left else -1.0
    x_tie = cx + side * m
    dash = (5.0, 4.0) if dashed else None

    ln = Line()
    ln.X1 = x_tie
    ln.Y1 = y_min
    ln.X2 = x_tie
    ln.Y2 = y_max
    ln.Stroke = br
    ln.StrokeThickness = th
    try:
        ln.StrokeStartLineCap = PenLineCap.Flat
        ln.StrokeEndLineCap = PenLineCap.Flat
    except Exception:
        pass
    if dash:
        try:
            from System.Windows.Media import DoubleCollection

            arr = DoubleCollection()
            for v in dash:
                arr.Add(float(v))
            ln.StrokeDashArray = arr
        except Exception:
            pass
    try:
        Panel.SetZIndex(ln, 22)
    except Exception:
        pass
    canv.Children.Add(ln)

    draw_hook_135_traba_end(
        canv, bx_top, by_top, m, tip_left, u"top", br, th, z_index=22, dash=dash,
    )
    draw_hook_135_traba_end(
        canv, bx_bot, by_bot, m, tip_left, u"bottom", br, th, z_index=22, dash=dash,
    )
    return ln
