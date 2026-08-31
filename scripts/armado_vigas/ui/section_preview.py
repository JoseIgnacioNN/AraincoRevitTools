# -*- coding: utf-8 -*-
"""Preview de sección transversal (confinamiento E, capas longitudinales)."""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Point, Thickness, TextAlignment, Size, FontWeights
from System.Windows.Controls import Canvas, TextBlock, Panel as WpfPanel
from System.Windows.Media import (
    DoubleCollection,
    PathFigure,
    LineSegment,
    ArcSegment,
    PathGeometry,
    PenLineJoin,
    RotateTransform,
    SweepDirection,
    StreamGeometry,
    EdgeMode,
    RenderOptions,
    Color,
    SolidColorBrush,
)
from System.Windows.Shapes import Ellipse, Line, Path, Rectangle

from armado_vigas.domain.confinement import (
    ensure_beam_confinement,
    find_confin_def,
    is_conf_draft_defined,
)
from armado_vigas.domain.constants import ESTRIBO_SPACING_DEFAULT_CENT, ESTRIBO_SPACING_DEFAULT_EXT
from armado_vigas.domain.laterales import (
    LATERALES_DIAM_DEFAULT,
)
from armado_vigas.domain.layers import (
    beam_n_capas_inf,
    beam_n_capas_sup,
    ensure_beam_layers,
    first_layer_bar_count,
    layer_keys,
)
from armado_vigas.domain.stirrups import parse_beam_section
from armado_vigas.ui import layout as lay
from armado_vigas.ui.net_ui import freeze_freezable
from armado_vigas.ui.stirrup_draw_135 import (
    COLOR_ESTRIBO,
    COLOR_TRABA,
    STIRRUP_THICK,
    TIE_THICK,
    draw_estribo_bbox,
    draw_traba_vertical_column,
    stirrup_margin_px,
)
from armado_vigas.ui.wpf_controls import brush_hex

PREVIEW_CANVAS_W = lay.SECTION_CTRL_WIDTH_PX
PREVIEW_CANVAS_H = 222.0

# Separación entre capas longitudinales en la sección (centroide a centroide).
SECTION_LAYER_CENTROID_GAP_MM = 25.0


_RUBBER_PAIR_PAD = 5.0


def _rubber_tag(canvas, text, color_hex, left, top):
    try:
        tag = TextBlock()
        tag.Text = text
        tag.FontSize = 10.0
        tag.FontWeight = FontWeights.SemiBold
        tag.Foreground = brush_hex(color_hex)
        tag.IsHitTestVisible = False
        try:
            WpfPanel.SetZIndex(tag, 40)
        except Exception:
            pass
        Canvas.SetLeft(tag, float(left))
        Canvas.SetTop(tag, float(top))
        canvas.Children.Add(tag)
    except Exception:
        pass


def _add_dashed_segment(canvas, x1, y1, x2, y2, stroke, thick=1.4, dash=(5.0, 3.5), z=30):
    """Segmento discontinuo."""
    try:
        ln = Line()
        ln.X1 = float(x1)
        ln.Y1 = float(y1)
        ln.X2 = float(x2)
        ln.Y2 = float(y2)
        ln.Stroke = stroke
        ln.StrokeThickness = float(thick)
        ln.IsHitTestVisible = False
        try:
            arr = DoubleCollection()
            for v in dash:
                if float(v) > 0:
                    arr.Add(float(v))
            if arr.Count >= 1:
                ln.StrokeDashArray = arr
        except Exception:
            pass
        try:
            WpfPanel.SetZIndex(ln, int(z))
        except Exception:
            pass
        canvas.Children.Add(ln)
    except Exception:
        pass


def _columns_in_marquee(first_sup, first_inf, n_row, ox, oy, cx, cy, pad=4.0):
    """Indices de columna cuyo eje vertical cruzan el marquee."""
    hit = []
    try:
        L = min(float(ox), float(cx)) - pad
        R = max(float(ox), float(cx)) + pad
        T = min(float(oy), float(cy))
        B = max(float(oy), float(cy))
    except Exception:
        return hit
    n = int(n_row or 0)
    for i in range(n):
        try:
            bx = float(first_sup[i]["x"])
            y0 = min(float(first_sup[i]["y"]), float(first_inf[i]["y"]))
            y1 = max(float(first_sup[i]["y"]), float(first_inf[i]["y"]))
        except Exception:
            continue
        if bx < L or bx > R:
            continue
        if B < y0 or T > y1:
            continue
        hit.append(i)
    return hit


def _draw_windows_marquee(canvas, ox, oy, cx, cy):
    """Rectangulo segmentado + relleno origen a punta cursor (estilo Windows)."""
    try:
        left = min(float(ox), float(cx))
        top = min(float(oy), float(cy))
        w = abs(float(cx) - float(ox))
        h = abs(float(cy) - float(oy))
    except Exception:
        return
    if w < 1.0 and h < 1.0:
        cr = 5.5
        cross = brush_hex(u"#fbbf24", 230)
        _add_dashed_segment(canvas, ox - cr, oy, ox + cr, oy, cross, 1.25, (8.0, 0), 34)
        _add_dashed_segment(canvas, ox, oy - cr, ox, oy + cr, cross, 1.25, (8.0, 0), 34)
        return
    try:
        rect = Rectangle()
        rect.Width = max(1.0, w)
        rect.Height = max(1.0, h)
        Canvas.SetLeft(rect, left)
        Canvas.SetTop(rect, top)
        rect.Stroke = brush_hex(u"#0ea5e9", 230)
        rect.StrokeThickness = 1.25
        try:
            rect.StrokeDashArray = DoubleCollection([4.0, 3.0])
        except Exception:
            pass
        try:
            rect.Fill = SolidColorBrush(Color.FromArgb(31, 14, 165, 233))
        except Exception:
            rect.Fill = brush_hex(u"#0ea5e9", 28)
        rect.IsHitTestVisible = False
        try:
            WpfPanel.SetZIndex(rect, 28)
        except Exception:
            pass
        canvas.Children.Add(rect)
    except Exception:
        pass
    cr = 5.0
    cross_o = brush_hex(u"#fbbf24", 230)
    cross_c = brush_hex(u"#67e8f9", 210)
    _add_dashed_segment(canvas, ox - cr, oy, ox + cr, oy, cross_o, 1.2, (8.0, 0), 34)
    _add_dashed_segment(canvas, ox, oy - cr, ox, oy + cr, cross_o, 1.2, (8.0, 0), 34)
    _add_dashed_segment(canvas, cx - cr, cy, cx + cr, cy, cross_c, 1.1, (8.0, 0), 34)
    _add_dashed_segment(canvas, cx, cy - cr, cx, cy + cr, cross_c, 1.1, (8.0, 0), 34)


def _draw_conf_rubber_band(
    canvas,
    first_sup,
    first_inf,
    n_row,
    pending_bar,
    hover_bar,
    cursor_xy,
    origin_xy,
    m_stir,
    thick_e,
    thick_t,
):
    """Marquee Windows: clic1 ancla, mover (sin mantener), clic2 cierra."""
    if pending_bar is None or not first_sup or not first_inf:
        return
    try:
        p = int(pending_bar)
    except Exception:
        return
    if p < 0 or p >= int(n_row):
        return
    try:
        asup = first_sup[p]
        ainf = first_inf[p]
    except Exception:
        return

    m_stir = float(m_stir if m_stir is not None else stirrup_margin_px())
    guide = brush_hex(u"#22d3ee", 230)
    x_a = (float(asup["x"]) + float(ainf["x"])) * 0.5
    y_sup = float(asup["y"])
    y_inf = float(ainf["y"])

    ox = oy = None
    if origin_xy is not None:
        try:
            ox = float(origin_xy[0])
            oy = float(origin_xy[1])
        except Exception:
            ox = oy = None
    if ox is None:
        ox = x_a
        oy = (y_sup + y_inf) * 0.5

    cx = cy = None
    if cursor_xy is not None:
        try:
            cx = float(cursor_xy[0])
            cy = float(cursor_xy[1])
        except Exception:
            cx = cy = None
    if cx is None:
        cx, cy = ox, oy

    _add_dashed_segment(
        canvas, x_a, y_sup, x_a, y_inf, brush_hex(u"#fb923c", 160), 1.55, (3.0, 3.0), 31,
    )
    for pt in (asup, ainf):
        try:
            ring = Ellipse()
            rr = 6.5
            ring.Width = rr * 2.0
            ring.Height = rr * 2.0
            Canvas.SetLeft(ring, float(pt["x"]) - rr)
            Canvas.SetTop(ring, float(pt["y"]) - rr)
            ring.Stroke = guide
            ring.StrokeThickness = 1.4
            ring.Fill = brush_hex(u"#22d3ee", 45)
            ring.IsHitTestVisible = False
            try:
                WpfPanel.SetZIndex(ring, 32)
            except Exception:
                pass
            canvas.Children.Add(ring)
        except Exception:
            pass

    _draw_windows_marquee(canvas, ox, oy, cx, cy)

    cols = _columns_in_marquee(first_sup, first_inf, n_row, ox, oy, cx, cy)
    for i in cols:
        try:
            bsup = first_sup[i]
            binf = first_inf[i]
            _add_dashed_segment(
                canvas,
                float(bsup["x"]), float(bsup["y"]),
                float(binf["x"]), float(binf["y"]),
                brush_hex(u"#34d399", 140), 1.4, (4.0, 3.0), 29,
            )
        except Exception:
            pass

    h = None
    if hover_bar is not None:
        try:
            h = int(hover_bar)
        except Exception:
            h = None
    if h is not None and (h < 0 or h >= int(n_row)):
        h = None

    if h is not None and h == p:
        try:
            draw_traba_vertical_column(
                canvas,
                asup["x"], asup["y"],
                ainf["x"], ainf["y"],
                margin=m_stir,
                brush=brush_hex(COLOR_TRABA, 200),
                thick=float(thick_t or TIE_THICK),
                dashed=True,
            )
        except Exception:
            pass
        mid_y = (y_sup + y_inf) * 0.5
        _rubber_tag(canvas, u"TRABA [{0}] · 2.o clic".format(p), u"#fdba74", x_a + 6, mid_y - 6)
        return

    if h is not None and h != p:
        a, b = (p, h) if p < h else (h, p)
        try:
            pts = [first_sup[a], first_sup[b], first_inf[a], first_inf[b]]
            hit = draw_estribo_bbox(
                canvas,
                pts,
                margin=m_stir,
                brush=brush_hex(COLOR_ESTRIBO, 200),
                thick=float(thick_e or STIRRUP_THICK),
                dashed=True,
            )
            tag_x = hit[0] + 4 if hit is not None else float(first_sup[a]["x"])
            tag_y = max(0.0, (hit[1] - 16) if hit is not None else float(first_sup[a]["y"]) - 16)
            _rubber_tag(
                canvas,
                u"ESTRIBO E({0}-{1}) · 2.o clic".format(a, b),
                u"#a5f3fc",
                tag_x,
                tag_y,
            )
        except Exception:
            pass
        return

    rx0 = min(ox, cx)
    ry0 = min(oy, cy)
    _rubber_tag(
        canvas,
        u"Preview · mueva cursor · 2.o clic cierra",
        u"#94a3b8",
        rx0 + 4,
        max(0.0, ry0 - 14),
    )



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
    return max(1.2, min(3.6, r))


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
    freeze_freezable(geom)
    path = Path()
    path.Data = geom
    path.Stroke = stroke_brush
    path.StrokeThickness = r_est
    path.StrokeLineJoin = PenLineJoin.Round
    path.Fill = brush_hex(u"#000000", 0)
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

    # Cotas: un StreamGeometry (ticks + ejes) en vez de 6× Line.
    tick = brush_hex(u"#64748b", 140)
    segs_py = [
        (ox - 5.0, oy, ox - 1.0, oy),
        (ox - 5.0, oy + sec_h, ox - 1.0, oy + sec_h),
        (ox - 3.0, oy, ox - 3.0, oy + sec_h),
        (ox, oy + sec_h + 3.0, ox + sec_w, oy + sec_h + 3.0),
        (ox, oy + sec_h + 1.0, ox, oy + sec_h + 5.0),
        (ox + sec_w, oy + sec_h + 1.0, ox + sec_w, oy + sec_h + 5.0),
    ]
    try:
        sg = StreamGeometry()
        ctx = sg.Open()
        for x1, y1, x2, y2 in segs_py:
            ctx.BeginFigure(Point(x1, y1), False, False)
            ctx.LineTo(Point(x2, y2), True, False)
        ctx.Close()
        freeze_freezable(sg)
        path = Path()
        path.Data = sg
        path.Stroke = tick
        path.StrokeThickness = 0.7
        path.Fill = None
        try:
            RenderOptions.SetEdgeMode(path, EdgeMode.Aliased)
            path.SnapsToDevicePixels = True
        except Exception:
            pass
        canvas.Children.Add(path)
    except Exception:
        for x1, y1, x2, y2 in segs_py:
            ln = Line()
            ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
            ln.Stroke = tick
            ln.StrokeThickness = 0.7
            canvas.Children.Add(ln)



def draw_section_preview(
    canvas,
    beam,
    role_label=None,
    laterales_enabled=False,
    n_laterales=1,
    diam_laterales=LATERALES_DIAM_DEFAULT,
    interactive=False,
    pending_bar=None,
    hover_bar=None,
    conf_draw_mode=u"draw",
    snap_r=None,
    out_geom=None,
    show_footer=True,
    session=None,
    beams=None,
    beam_idx=None,
    preferred_tramo_sup=None,
    preferred_tramo_inf=None,
    preferred_tramo_sup_id=None,
    preferred_tramo_inf_id=None,
    cursor_xy=None,
    origin_xy=None,
):
    """Pinta sección (mismo layout en SUP/INF/LAT/CONF).

    ``interactive`` solo añade capacidad de dibujo CONF (snap, rubber-band, hits);
    no cambia geometría ni estilo del hormigón/barras/conf (estilo SUP).

    ``session`` / ``beams``: resuelven n/ø reales desde tramos SUP e INF
    (``tramo_armado``) para que CONF dibuje la misma cantidad de barras
    configurada en esas pestañas.
    """
    if canvas is None:
        return u""
    canvas.Children.Clear()
    if not beam:
        return u""
    try:
        RenderOptions.SetEdgeMode(canvas, EdgeMode.Aliased)
        canvas.SnapsToDevicePixels = True
    except Exception:
        pass

    ensure_beam_layers(beam)
    ensure_beam_confinement(beam)

    # Vista de pintura: n/ø de tramos SUP+INF (no defaults del dict viga).
    paint = beam
    try:
        from armado_vigas.domain.tramo_armado import (
            beam_section_arm_for_preview,
            sync_resolved_arm_onto_beam,
        )

        if session is None:
            try:
                from armado_vigas.revit.session import SESSION as _S

                session = _S
            except Exception:
                session = None

        resolved = beam_section_arm_for_preview(
            session,
            beam,
            beams=beams,
            beam_idx=beam_idx,
            preferred_tramo_sup=preferred_tramo_sup,
            preferred_tramo_inf=preferred_tramo_inf,
            preferred_tramo_sup_id=preferred_tramo_sup_id,
            preferred_tramo_inf_id=preferred_tramo_inf_id,
        )
        if resolved:
            paint = resolved
            # Alinear beam de dominio con n/ø reales (columnas conf. E/T).
            sync_resolved_arm_onto_beam(beam, paint)
            ensure_beam_confinement(beam)
    except Exception:
        paint = beam

    # Conf (draft) sigue en el beam de dominio; capas en paint.
    cw, ch = _canvas_dims(canvas)
    w_cm, h_cm = parse_beam_section(paint.get("type") or beam.get("type"))
    # Layout fijo (como pestaña SUP) en las 4 pestañas.
    pad_x = 10.0
    pad_top = 4.0
    label_h = 12.0
    ox, oy, sec_w, sec_h = _fit_section_rect(
        w_cm, h_cm, cw, ch, pad_x=pad_x, pad_top=pad_top, label_h=label_h,
    )

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
    inner.Fill = brush_hex(u"#000000", 0)
    canvas.Children.Add(inner)

    bar_pad = max(2.5, cover * 0.55)
    bar_x0 = st_x + bar_pad
    bar_x1 = st_x + st_w - bar_pad
    n_capas_sup = beam_n_capas_sup(paint)
    n_capas_inf = beam_n_capas_inf(paint)
    # scale = px / cm → 25 mm entre centroides de capas SUP e INF.
    layer_step = (
        (float(SECTION_LAYER_CENTROID_GAP_MM) / 10.0) * float(scale)
        if max(n_capas_sup, n_capas_inf) > 1
        else 0.0
    )

    layer_colors = [
        (u"#22d3ee", u"#0891b2", u"#f87171", u"#b91c1c", 255),
        (u"#38bdf8", u"#0284c7", u"#fb7185", u"#e11d48", 224),
        (u"#7dd3fc", u"#0369a1", u"#fda4af", u"#be123c", 184),
    ]

    k1 = layer_keys(1)
    n_first = first_layer_bar_count(paint)
    n_row = int(n_first)

    sup_y1 = st_y + bar_pad
    inf_y1 = st_y + st_h - bar_pad
    first_sup = _distribute_points(n_row, bar_x0, bar_x1, sup_y1)
    first_inf = _distribute_points(n_row, bar_x0, bar_x1, inf_y1)

    # Capas 2+ (estilo SUP en todas las pestañas)
    for layer_num in range(2, n_capas_sup + 1):
        k = layer_keys(layer_num)
        n_s = int(paint.get(k["nSup"]) or 2)
        d_s = int(paint.get(k["diamSup"]) or 16)
        sup_y = st_y + bar_pad + (layer_num - 1) * layer_step
        sup_pts = _distribute_points(n_s, bar_x0, bar_x1, sup_y)
        ci = min(layer_num - 1, len(layer_colors) - 1)
        cs, css, _, _, op = layer_colors[ci]
        max_r_sup = (bar_x1 - bar_x0) / max(1, n_s - 1) * 0.36 if n_s > 1 else (bar_x1 - bar_x0) * 0.14
        r_s = _bar_radius(d_s, scale, max_r_sup) * 0.92
        for pt in sup_pts:
            _add_dot(canvas, pt["x"], pt["y"], r_s, cs, css, op)

    for layer_num in range(2, n_capas_inf + 1):
        k = layer_keys(layer_num)
        n_i = int(paint.get(k["nInf"]) or 2)
        d_i = int(paint.get(k["diamInf"]) or 16)
        inf_y = st_y + st_h - bar_pad - (layer_num - 1) * layer_step
        inf_pts = _distribute_points(n_i, bar_x0, bar_x1, inf_y)
        ci = min(layer_num - 1, len(layer_colors) - 1)
        _, _, ci_fill, cis, op = layer_colors[ci]
        max_r_inf = (bar_x1 - bar_x0) / max(1, n_i - 1) * 0.36 if n_i > 1 else (bar_x1 - bar_x0) * 0.14
        r_i = _bar_radius(d_i, scale, max_r_inf) * 0.92
        for pt in inf_pts:
            _add_dot(canvas, pt["x"], pt["y"], r_i, ci_fill, cis, op)

    k1_d_s = int(paint.get(k1["diamSup"]) or 16)
    k1_d_i = int(paint.get(k1["diamInf"]) or 16)
    max_r1 = (bar_x1 - bar_x0) / max(1, n_row - 1) * 0.36 if n_row > 1 else (bar_x1 - bar_x0) * 0.14
    r_s1 = _bar_radius(k1_d_s, scale, max_r1)
    r_i1 = _bar_radius(k1_d_i, scale, max_r1)

    # Solo geometría ya definida (pairs / ties / peri). Sin draft → no E/T.
    if is_conf_draft_defined(beam):
        conf = find_confin_def(beam)
    else:
        conf = {
            u"label": u"Dibujo libre",
            u"perimetral": False,
            u"pairs": [],
            u"ties": [],
        }
    m_stir = stirrup_margin_px()
    thick_e = STIRRUP_THICK
    thick_t = TIE_THICK

    if conf.get("perimetral") and first_sup and first_inf:
        try:
            draw_estribo_bbox(
                canvas, list(first_sup) + list(first_inf),
                margin=m_stir, brush=brush_hex(COLOR_ESTRIBO), thick=thick_e,
            )
        except Exception:
            pass

    pair_hits = []
    for pair in conf.get("pairs") or []:
        if len(pair) < 2:
            continue
        i0, i1 = int(pair[0]), int(pair[1])
        if i0 >= n_row or i1 >= n_row:
            continue
        a, b = min(i0, i1), max(i0, i1)
        pts = [first_sup[a], first_sup[b], first_inf[a], first_inf[b]]
        try:
            hit = draw_estribo_bbox(
                canvas, pts, margin=m_stir,
                brush=brush_hex(COLOR_ESTRIBO), thick=thick_e,
            )
        except Exception:
            hit = None
        if hit is not None:
            pair_hits.append({u"i0": a, u"i1": b, u"rx": hit[0], u"ry": hit[1], u"rw": hit[2], u"rh": hit[3]})

    for idx in conf.get("ties") or []:
        ti = int(idx)
        if ti >= n_row:
            continue
        try:
            draw_traba_vertical_column(
                canvas,
                first_sup[ti]["x"], first_sup[ti]["y"],
                first_inf[ti]["x"], first_inf[ti]["y"],
                margin=m_stir, brush=brush_hex(COLOR_TRABA), thick=thick_t,
            )
        except Exception:
            pass

    # Barras 1ª capa (mismo estilo SUP; halo snap ≈ barra + margen pequeño).
    # Radio de hit: ligeramente mayor que el halo (no el fijo 22–36 px).
    r_bar_max = max(float(r_s1 or 1.0), float(r_i1 or 1.0))
    snap_hit_r = max(r_bar_max * 1.75 + 2.0, r_bar_max + 4.0)
    if snap_r is not None:
        try:
            # Solo permite ampliar un poco (p. ej. zoom); no forzar anillos enormes.
            snap_hit_r = max(snap_hit_r, min(float(snap_r), r_bar_max * 2.6 + 3.0))
        except Exception:
            pass
    hits = []
    for i in range(n_row):
        is_a = interactive and pending_bar is not None and int(pending_bar) == i
        is_h = interactive and hover_bar is not None and int(hover_bar) == i
        r_sup = r_s1 * (1.08 if is_a or is_h else 1.0)
        r_inf = r_i1 * (1.08 if is_a or is_h else 1.0)
        cs = u"#a5f3fc" if is_a else (u"#67e8f9" if is_h else u"#22d3ee")
        ci = u"#fda4af" if is_a else (u"#fb7185" if is_h else u"#f87171")
        if interactive and (conf_draw_mode in (u"draw", u"erase") or is_a or is_h):
            for pt, bar_r in (
                (first_sup[i], r_s1),
                (first_inf[i], r_i1),
            ):
                # Halo visual: solo un poco más grande que la barra.
                halo_r = max(float(bar_r) + 1.5, float(bar_r) * 1.22)
                if is_a or is_h:
                    halo_r = max(halo_r, float(bar_r) + 2.0)
                ring = Ellipse()
                d = halo_r * 2.0
                ring.Width = d
                ring.Height = d
                Canvas.SetLeft(ring, pt["x"] - halo_r)
                Canvas.SetTop(ring, pt["y"] - halo_r)
                ring.Stroke = brush_hex(u"#22d3ee", 180 if (is_h or is_a) else 70)
                ring.StrokeThickness = 1.1 if (is_h or is_a) else 0.75
                if not (is_h or is_a):
                    try:
                        ring.StrokeDashArray = DoubleCollection([2.0, 2.0])
                    except Exception:
                        pass
                ring.Fill = brush_hex(u"#000000", 0)
                ring.IsHitTestVisible = False
                canvas.Children.Add(ring)

        _add_dot(canvas, first_sup[i]["x"], first_sup[i]["y"], r_sup, cs, u"#0891b2", 255)
        _add_dot(canvas, first_inf[i]["x"], first_inf[i]["y"], r_inf, ci, u"#b91c1c", 255)
        hits.append({u"i": i, u"x": first_sup[i]["x"], u"y": first_sup[i]["y"]})
        hits.append({u"i": i, u"x": first_inf[i]["x"], u"y": first_inf[i]["y"]})
        # Índices de columna (0,1,…) no en CONF: el snap basta sin etiquetas.

    # Rubber-band encima de barras: guía entre clic 1 y 2.
    if interactive and conf_draw_mode == u"draw" and pending_bar is not None:
        _draw_conf_rubber_band(
            canvas,
            first_sup,
            first_inf,
            n_row,
            pending_bar,
            hover_bar,
            cursor_xy,
            origin_xy,
            m_stir,
            thick_e,
            thick_t,
        )

    # Laterales iguales en las 4 pestañas: X = columnas extremas SUP/INF.
    if laterales_enabled:
        edge_xs = None
        if first_sup and len(first_sup) >= 1:
            edge_xs = (float(first_sup[0]["x"]), float(first_sup[-1]["x"]))
        _draw_lateral_preview_dots(
            canvas,
            paint,
            st_x,
            st_y,
            st_w,
            st_h,
            bar_pad,
            layer_step,
            n_capas_sup,
            n_capas_inf,
            scale,
            int(n_laterales or 0),
            int(diam_laterales or LATERALES_DIAM_DEFAULT),
            edge_xs=edge_xs,
            outer_y=oy,
            outer_h=sec_h,
        )

    # Cotas b/h: solo en SUP/INF/LAT (preview estática). En CONF sobran.
    if not interactive:
        _add_dim_labels(canvas, ox, oy, sec_w, sec_h, w_cm, h_cm)

    if show_footer:
        est_ext = beam.get("estExtDiam") or paint.get("estExtDiam") or 10
        est_cent = beam.get("estCentDiam") or paint.get("estCentDiam") or 8
        foot = TextBlock()
        nE = len(conf.get(u"pairs") or [])
        nT = len(conf.get(u"ties") or [])
        peri = bool(conf.get(u"perimetral"))
        if interactive:
            if nE or nT or peri:
                foot.Text = u"1.er clic · mueva · 2.º clic · {0}E/{1}T{2} · {3}ø1ª".format(
                    nE, nT, u"+P" if peri else u"", n_row,
                )
            else:
                foot.Text = u"Sin conf. · {0} barras 1ª · 1.er clic ancla · mueva · 2.º clic".format(
                    n_row,
                )
            foot.Foreground = brush_hex(u"#22d3ee")
        else:
            if nE or nT or peri:
                foot.Text = u"Conf. {0}E/{1}T{2} · Ext ø{3}@{4} · Cent ø{5}@{6}".format(
                    nE, nT, u"+P" if peri else u"",
                    est_ext, beam.get("estExtSpacing") or ESTRIBO_SPACING_DEFAULT_EXT,
                    est_cent, beam.get("estCentSpacing") or ESTRIBO_SPACING_DEFAULT_CENT,
                )
            else:
                foot.Text = u"Sin conf. E · Ext ø{0}@{1} · Cent ø{2}@{3}".format(
                    est_ext, beam.get("estExtSpacing") or ESTRIBO_SPACING_DEFAULT_EXT,
                    est_cent, beam.get("estCentSpacing") or ESTRIBO_SPACING_DEFAULT_CENT,
                )
            foot.Foreground = brush_hex(u"#94a3b8")
        foot.FontSize = 7.5
        foot.Width = cw
        foot.TextAlignment = TextAlignment.Center
        foot.IsHitTestVisible = False
        Canvas.SetLeft(foot, 0.0)
        Canvas.SetTop(foot, ch - 13.0)
        canvas.Children.Add(foot)

    if out_geom is not None:
        out_geom.clear()
        out_geom.update({
            u"ox": ox, u"oy": oy, u"secW": sec_w, u"secH": sec_h,
            u"hits": hits, u"n": n_row, u"snapR": snap_hit_r,
            u"first_sup": first_sup, u"first_inf": first_inf,
            u"pairHits": pair_hits,
            u"nSup": int(paint.get(u"nSup") or n_row),
            u"nInf": int(paint.get(u"nInf") or n_row),
            u"diamSup": k1_d_s,
            u"diamInf": k1_d_i,
        })

    n_conf = first_layer_bar_count(paint)
    lbl = conf.get("label") or u"Dibujo libre"
    short = lbl if len(lbl) <= 32 else lbl[:30] + u"…"
    role = role_label or u"Sección"
    return u"{0} · {1} · {2}b · {3}".format(
        beam.get("id") or paint.get("id") or u"?",
        beam.get("type") or paint.get("type") or u"?",
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
    edge_xs=None,
    outer_y=None,
    outer_h=None,
):
    """Laterales en alma: X = columnas extremas SUP/INF;
    Y = regla face_clear ``H/ceil(H/200)`` simétrica desde caras de sección.
    """
    try:
        n_lat = int(n_lat or 0)
    except Exception:
        n_lat = 0
    if n_lat < 1:
        return

    from armado_vigas.domain.laterales import (
        beam_section_height_mm,
        lateral_face_clear_mm,
        lateral_ys_from_face_mm,
    )
    from armado_vigas.domain.stirrups import section_height_mm

    # Alineamiento vertical (misma X) con barras longitudinales extremas.
    x_faces = []
    if edge_xs is not None:
        try:
            x0 = float(edge_xs[0])
            x1 = float(edge_xs[1] if len(edge_xs) > 1 else edge_xs[0])
            x_faces = [x0] if abs(x1 - x0) < 0.5 else [x0, x1]
        except Exception:
            x_faces = []
    if not x_faces:
        # Fallback: extremos del reparto longitudinal (mismo bar_pad que SUP/INF).
        x_left = float(st_x) + float(bar_pad)
        x_right = float(st_x) + float(st_w) - float(bar_pad)
        x_faces = [x_left] if abs(x_right - x_left) < 0.5 else [x_left, x_right]

    try:
        h_mm = float(beam_section_height_mm(beam) or 0.0)
    except Exception:
        h_mm = 0.0
    if h_mm <= 0:
        try:
            h_mm = float(section_height_mm((beam or {}).get("type")))
        except Exception:
            h_mm = 0.0
    if h_mm <= 0:
        # Sin H: no forzar posiciones erróneas.
        return

    y0 = float(outer_y) if outer_y is not None else float(st_y)
    oh = float(outer_h) if outer_h is not None else float(st_h)
    # scale = px / cm → mm:  * (scale/10)
    mm_to_px = float(scale) / 10.0
    ys_mm = lateral_ys_from_face_mm(h_mm, n_lat)
    if not ys_mm:
        return

    face_clear = lateral_face_clear_mm(h_mm)
    span_px = max(
        0.5,
        (float(h_mm) - 2.0 * float(face_clear)) * mm_to_px,
    )
    r = _bar_radius(diam_mm, scale, max(1.5, span_px / max(1, n_lat) * 0.35))
    fill = u"#c4b5fd"
    stroke = u"#8b5cf6"
    for x_face in x_faces:
        for ymm in ys_mm:
            y = y0 + float(ymm) * mm_to_px
            # Clamp suave al rectángulo de sección si hubo desfase numérico.
            y = max(y0 + 0.5, min(y0 + oh - 0.5, y))
            _add_dot(canvas, x_face, y, r, fill, stroke, 220)


def _add_dot(canvas, x, y, r, fill_hex, stroke_hex, alpha=255, hit=False):
    el = Ellipse()
    d = r * 2.0
    el.Width = d
    el.Height = d
    Canvas.SetLeft(el, x - r)
    Canvas.SetTop(el, y - r)
    el.Fill = brush_hex(fill_hex, alpha)
    el.Stroke = brush_hex(stroke_hex, min(255, alpha + 20))
    el.StrokeThickness = 0.9
    if not hit:
        try:
            el.IsHitTestVisible = False
        except Exception:
            pass
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
        beam.get("estExtSpacing") or ESTRIBO_SPACING_DEFAULT_EXT,
        beam.get("estCentDiam") or 8,
        beam.get("estCentSpacing") or ESTRIBO_SPACING_DEFAULT_CENT,
    )
