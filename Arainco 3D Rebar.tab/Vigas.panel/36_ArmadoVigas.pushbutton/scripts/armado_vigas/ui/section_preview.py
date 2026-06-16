# -*- coding: utf-8 -*-
"""Preview de sección transversal (confinamiento E, capas longitudinales)."""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Point, Thickness, TextAlignment, Size
from System.Windows.Controls import Canvas, TextBlock
from System.Windows.Media import (
    DoubleCollection,
    PathFigure,
    LineSegment,
    ArcSegment,
    PathGeometry,
    PenLineJoin,
    RotateTransform,
    SolidColorBrush,
    Color,
    SweepDirection,
)
from System.Windows.Shapes import Ellipse, Line, Path, Rectangle

from armado_vigas.domain.confinement import ensure_beam_confinement, find_confin_def
from armado_vigas.domain.laterales import conf_diam_mm, lateral_clear_mm
from armado_vigas.domain.layers import (
    beam_n_capas_inf,
    beam_n_capas_sup,
    ensure_beam_layers,
    first_layer_bar_count,
    layer_keys,
)
from armado_vigas.domain.stirrups import parse_beam_section
from armado_vigas.ui import layout as lay
from armado_vigas.ui.wpf_controls import brush_hex

PREVIEW_CANVAS_W = lay.SECTION_CTRL_WIDTH_PX
PREVIEW_CANVAS_H = 222.0


def _canvas_dims(canvas):
    try:
        w = float(canvas.Width)
        if w > 1.0:
            cw = w
        else:
            cw = PREVIEW_CANVAS_W
    except Exception:
        cw = PREVIEW_CANVAS_W
    try:
        h = float(canvas.Height)
        if h > 1.0:
            ch = h
        else:
            ch = PREVIEW_CANVAS_H
    except Exception:
        ch = PREVIEW_CANVAS_H
    return cw, ch


def _fit_section_rect(w_cm, h_cm, canvas_w, canvas_h, pad_x=10.0, pad_top=4.0, label_h=12.0):
    max_w = canvas_w - pad_x * 2.0
    max_h = canvas_h - pad_top - label_h - 4.0
    b = float(w_cm)
    h = float(h_cm)
    if h >= b:
        sec_h = max_h
        sec_w = sec_h * (b / h)
        if sec_w > max_w:
            sec_w = max_w
            sec_h = sec_w * (h / b)
    else:
        sec_w = max_w
        sec_h = sec_w * (h / b)
        if sec_h > max_h:
            sec_h = max_h
            sec_w = sec_h * (b / h)
    ox = (canvas_w - sec_w) * 0.5
    oy = pad_top + (max_h - sec_h) * 0.5
    return ox, oy, sec_w, sec_h


def _distribute_points(count, x0, x1, y):
    if count <= 0:
        return []
    if count == 1:
        return [{"x": (x0 + x1) * 0.5, "y": y}]
    step = (x1 - x0) / float(count - 1)
    return [{"x": x0 + i * step, "y": y} for i in range(count)]


def _distribute_points_vertical(count, y0, y1, x):
    if count <= 0:
        return []
    if count == 1:
        return [{"x": x, "y": (y0 + y1) * 0.5}]
    step = (y1 - y0) / float(count - 1)
    return [{"x": x, "y": y0 + i * step} for i in range(count)]


def _bar_radius(diam_mm, scale, max_from_spacing=None):
    r = (float(diam_mm) / 18.0) * 1.45 * scale
    if max_from_spacing is not None and max_from_spacing > 0:
        r = min(r, max_from_spacing)
    return max(1.0, min(2.2, r))


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


def _add_dim_labels(canvas, ox, oy, sec_w, sec_h, w_cm, h_cm):
    dim_brush = brush_hex(u"#64748b")
    dim_size = 8.0

    h_lbl = TextBlock()
    h_lbl.Text = u"{0:.0f}".format(float(w_cm)) if float(w_cm) == int(w_cm) else u"{0:.1f}".format(float(w_cm))
    h_lbl.FontSize = dim_size
    h_lbl.Foreground = dim_brush
    Canvas.SetLeft(h_lbl, ox + sec_w * 0.5 - 8.0)
    Canvas.SetTop(h_lbl, oy + sec_h + 6.0)
    canvas.Children.Add(h_lbl)

    v_lbl = TextBlock()
    v_lbl.Text = u"{0:.0f}".format(float(h_cm)) if float(h_cm) == int(h_cm) else u"{0:.1f}".format(float(h_cm))
    v_lbl.FontSize = dim_size
    v_lbl.Foreground = dim_brush
    v_lbl.RenderTransformOrigin = Point(0.5, 0.5)
    v_lbl.RenderTransform = RotateTransform(-90.0)
    Canvas.SetLeft(v_lbl, ox - 18.0)
    Canvas.SetTop(v_lbl, oy + sec_h * 0.5 - 4.0)
    canvas.Children.Add(v_lbl)

    # Cotas mínimas
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
    for x_pos in (ox, ox + sec_w):
        ln = Line()
        ln.X1 = x_pos
        ln.Y1 = oy + sec_h + 1.0
        ln.X2 = x_pos
        ln.Y2 = oy + sec_h + 5.0
        ln.Stroke = tick
        ln.StrokeThickness = 0.7
        canvas.Children.Add(ln)


def draw_section_preview(
    canvas,
    beam,
    role_label=None,
    laterales_enabled=False,
    n_laterales=1,
    diam_laterales=16,
):
    if canvas is None:
        return u""
    canvas.Children.Clear()
    if not beam:
        return u""

    ensure_beam_layers(beam)
    ensure_beam_confinement(beam)
    cw, ch = _canvas_dims(canvas)
    w_cm, h_cm = parse_beam_section(beam.get("type"))
    ox, oy, sec_w, sec_h = _fit_section_rect(w_cm, h_cm, cw, ch)

    cover = max(3.5, min(sec_w, sec_h) * 0.11)
    inner_x = ox + cover
    inner_y = oy + cover
    inner_w = sec_w - cover * 2.0
    inner_h = sec_h - cover * 2.0
    st_inset = 2.2
    st_x = inner_x + st_inset
    st_y = inner_y + st_inset
    st_w = inner_w - st_inset * 2.0
    st_h = inner_h - st_inset * 2.0
    scale = sec_h / max(float(h_cm), 1.0)

    outer = Rectangle()
    outer.Width = sec_w
    outer.Height = sec_h
    Canvas.SetLeft(outer, ox)
    Canvas.SetTop(outer, oy)
    outer.RadiusX = 1.8
    outer.RadiusY = 1.8
    outer.Stroke = brush_hex(u"#5bb8d4")
    outer.StrokeThickness = 1.3
    outer.Fill = brush_hex(u"#0a1620", 220)
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

    bar_pad = max(2.5, cover * 0.55)
    bar_x0 = st_x + bar_pad
    bar_x1 = st_x + st_w - bar_pad
    n_capas_sup = beam_n_capas_sup(beam)
    n_capas_inf = beam_n_capas_inf(beam)
    n_capas = max(n_capas_sup, n_capas_inf)
    layer_step = (
        max(3.5, min(st_h * 0.15, (st_h - bar_pad * 2.0) / max(1.0, n_capas * 2.0)))
        if n_capas > 1
        else 0.0
    )

    layer_colors = [
        (u"#22d3ee", u"#0891b2", u"#f87171", u"#b91c1c", 255),
        (u"#38bdf8", u"#0284c7", u"#fb7185", u"#e11d48", 224),
        (u"#7dd3fc", u"#0369a1", u"#fda4af", u"#be123c", 184),
    ]

    k1 = layer_keys(1)
    n_s1 = int(beam.get(k1["nSup"]) or 2)
    n_i1 = int(beam.get(k1["nInf"]) or 2)
    sup_y1 = st_y + bar_pad
    inf_y1 = st_y + st_h - bar_pad
    first_sup = _distribute_points(n_s1, bar_x0, bar_x1, sup_y1)
    first_inf = _distribute_points(n_i1, bar_x0, bar_x1, inf_y1)

    for layer_num in range(1, n_capas_sup + 1):
        k = layer_keys(layer_num)
        n_s = int(beam.get(k["nSup"]) or 2)
        d_s = int(beam.get(k["diamSup"]) or 16)
        sup_y = st_y + bar_pad + (layer_num - 1) * layer_step
        sup_pts = _distribute_points(n_s, bar_x0, bar_x1, sup_y)
        ci = min(layer_num - 1, len(layer_colors) - 1)
        cs, css, _, _, op = layer_colors[ci]
        max_r_sup = (bar_x1 - bar_x0) / max(1, n_s - 1) * 0.36 if n_s > 1 else (bar_x1 - bar_x0) * 0.14
        r_s = _bar_radius(d_s, scale, max_r_sup) * (1.0 if layer_num == 1 else 0.92)
        for pt in sup_pts:
            _add_dot(canvas, pt["x"], pt["y"], r_s, cs, css, op)

    for layer_num in range(1, n_capas_inf + 1):
        k = layer_keys(layer_num)
        n_i = int(beam.get(k["nInf"]) or 2)
        d_i = int(beam.get(k["diamInf"]) or 16)
        inf_y = st_y + st_h - bar_pad - (layer_num - 1) * layer_step
        inf_pts = _distribute_points(n_i, bar_x0, bar_x1, inf_y)
        ci = min(layer_num - 1, len(layer_colors) - 1)
        _, _, ci_fill, cis, op = layer_colors[ci]
        max_r_inf = (bar_x1 - bar_x0) / max(1, n_i - 1) * 0.36 if n_i > 1 else (bar_x1 - bar_x0) * 0.14
        r_i = _bar_radius(d_i, scale, max_r_inf) * (1.0 if layer_num == 1 else 0.92)
        for pt in inf_pts:
            _add_dot(canvas, pt["x"], pt["y"], r_i, ci_fill, cis, op)

    conf = find_confin_def(beam)
    r_est = max(0.9, float(beam.get("estCentDiam") or beam.get("estExtDiam") or 8) / 9.0)
    stir_stroke = brush_hex(u"#34d399")

    if conf.get("perimetral"):
        _add_outer_stirrup(canvas, st_x, st_y, st_w, st_h, stir_stroke, r_est)

    pad = max(2.8, r_est * 0.85)
    for pair in conf.get("pairs") or []:
        if len(pair) < 2:
            continue
        i0, i1 = pair[0], pair[1]
        if i0 >= len(first_sup) or i1 >= len(first_sup):
            continue
        if i0 >= len(first_inf) or i1 >= len(first_inf):
            continue
        xs = [first_sup[i0]["x"], first_sup[i1]["x"], first_inf[i0]["x"], first_inf[i1]["x"]]
        ys = [first_sup[i0]["y"], first_sup[i1]["y"], first_inf[i0]["y"], first_inf[i1]["y"]]
        rx = min(xs) - pad
        ry = min(ys) - pad
        rw = max(xs) - min(xs) + pad * 2.0
        rh = max(ys) - min(ys) + pad * 2.0
        if rw < 3.0 or rh < 3.0:
            continue
        rect = Rectangle()
        rect.Width = rw
        rect.Height = rh
        Canvas.SetLeft(rect, rx)
        Canvas.SetTop(rect, ry)
        rect.Stroke = brush_hex(u"#4ade80")
        rect.StrokeThickness = max(0.7, r_est * 0.55)
        rect.RadiusX = 1.3
        rect.RadiusY = 1.3
        rect.Fill = SolidColorBrush(Color.FromArgb(10, 74, 222, 128))
        canvas.Children.Add(rect)

    for idx in conf.get("ties") or []:
        if idx >= len(first_sup) or idx >= len(first_inf):
            continue
        x = (first_sup[idx]["x"] + first_inf[idx]["x"]) * 0.5
        ln = Line()
        ln.X1 = x
        ln.Y1 = first_sup[idx]["y"]
        ln.X2 = x
        ln.Y2 = first_inf[idx]["y"]
        ln.Stroke = brush_hex(u"#fb923c")
        ln.StrokeThickness = max(0.7, r_est * 0.55)
        canvas.Children.Add(ln)

    if laterales_enabled:
        _draw_lateral_preview_dots(
            canvas,
            beam,
            st_x,
            st_y,
            st_w,
            st_h,
            bar_pad,
            layer_step,
            n_capas_sup,
            n_capas_inf,
            scale,
            int(n_laterales or 1),
            int(diam_laterales or 16),
        )

    _add_dim_labels(canvas, ox, oy, sec_w, sec_h, w_cm, h_cm)

    est_ext = beam.get("estExtDiam") or 10
    est_cent = beam.get("estCentDiam") or 8
    foot = TextBlock()
    foot.Text = u"Ext ø{0}@{1} · Cent ø{2}@{3}".format(
        est_ext, beam.get("estExtSpacing") or 125,
        est_cent, beam.get("estCentSpacing") or 200,
    )
    foot.FontSize = 7.5
    foot.Foreground = brush_hex(u"#94a3b8")
    foot.Width = cw
    foot.TextAlignment = TextAlignment.Center
    Canvas.SetLeft(foot, 0.0)
    Canvas.SetTop(foot, ch - 13.0)
    canvas.Children.Add(foot)

    n_conf = first_layer_bar_count(beam)
    lbl = conf.get("label") or u"Perimetral"
    short = lbl if len(lbl) <= 32 else lbl[:30] + u"…"
    role = role_label or u"Cent / confin."
    return u"{0} · {1} · {2}b · {3}".format(
        beam.get("id") or u"?",
        beam.get("type") or u"?",
        n_conf,
        short,
    )


def _draw_lateral_preview_dots(
    canvas,
    beam,
    st_x,
    st_y,
    st_w,
    st_h,
    bar_pad,
    layer_step,
    n_capas_sup,
    n_capas_inf,
    scale,
    n_lat,
    diam_mm,
):
    """Puntos en caras del alma (izq/der), zona entre fibras flexión + clear."""
    if n_lat < 1:
        return
    conf_d = conf_diam_mm(beam)
    inset = max(2.0, float(conf_d) / 18.0 * 1.35 * scale)
    x_left = st_x + inset
    x_right = st_x + st_w - inset
    clear_px = max(3.0, lateral_clear_mm(beam) * scale / 10.0)
    y_top = st_y + bar_pad + max(0.0, n_capas_sup - 1) * layer_step + clear_px
    y_bot = st_y + st_h - bar_pad - max(0.0, n_capas_inf - 1) * layer_step - clear_px
    if y_bot <= y_top + 2.0:
        y_top = st_y + st_h * 0.28
        y_bot = st_y + st_h * 0.72
    r = _bar_radius(diam_mm, scale, max(1.5, (y_bot - y_top) / max(1, n_lat) * 0.35))
    fill = u"#c4b5fd"
    stroke = u"#8b5cf6"
    for x_face in (x_left, x_right):
        for pt in _distribute_points_vertical(n_lat, y_top, y_bot, x_face):
            _add_dot(canvas, pt["x"], pt["y"], r, fill, stroke, 220)


def _add_dot(canvas, x, y, r, fill_hex, stroke_hex, alpha=255):
    el = Ellipse()
    d = r * 2.0
    el.Width = d
    el.Height = d
    Canvas.SetLeft(el, x - r)
    Canvas.SetTop(el, y - r)
    el.Fill = brush_hex(fill_hex, alpha)
    el.Stroke = brush_hex(stroke_hex, min(255, alpha + 20))
    el.StrokeThickness = 0.9
    canvas.Children.Add(el)


def section_meta_lines(beam, role_label=None):
    from armado_vigas.domain.stirrups import compute_stirrup_zones
    from armado_vigas.domain.layers import layer_keys

    ensure_beam_layers(beam)
    ensure_beam_confinement(beam)
    n_capas_sup = beam_n_capas_sup(beam)
    n_capas_inf = beam_n_capas_inf(beam)
    sup_parts = []
    for layer_num in range(1, n_capas_sup + 1):
        k = layer_keys(layer_num)
        sup_parts.append(
            u"{0} {1}ø{2}".format(
                k["label"],
                beam.get(k["nSup"]) or 2,
                beam.get(k["diamSup"]) or 16,
            )
        )
    inf_parts = []
    for layer_num in range(1, n_capas_inf + 1):
        k = layer_keys(layer_num)
        inf_parts.append(
            u"{0} {1}ø{2}".format(
                k["label"],
                beam.get(k["nInf"]) or 2,
                beam.get(k["diamInf"]) or 16,
            )
        )
    cap_txt = u"Sup {0} · Inf {1}".format(
        u" · ".join(sup_parts) if sup_parts else u"—",
        u" · ".join(inf_parts) if inf_parts else u"—",
    )
    plan = compute_stirrup_zones(beam)
    if plan.get("mode") == "single":
        z = (plan.get("zones") or [{}])[0]
        stir = u"1 lote · L {0} mm".format(z.get("lenMm") or 0)
    else:
        stir = u"Ext ×2 · Cent {0} mm".format(plan.get("L_cent") or 0)
    role = role_label or u"Cent / confin."
    return u"{0}\n{1} · {2}\nExt ø{3} @ {4} · Cent ø{5} @ {6}".format(
        cap_txt,
        stir,
        role,
        beam.get("estExtDiam") or 10,
        beam.get("estExtSpacing") or 125,
        beam.get("estCentDiam") or 8,
        beam.get("estCentSpacing") or 200,
    )
