# -*- coding: utf-8 -*-
"""Canvas de elevación: stack de muros (estilo Muros V2) + ciclo Auto/Tramo/Cont.

Dibujo fiel a Armado Muros V2:
- Escala uniforme px/mm (mismo para ancho y alto) → ``draw_w/draw_h ≈ L/H``.
- Fit vertical como máximo 8 muros; con más, escala fija y scrollbar (~8 visibles).
- Bloque (fustes + niveles + Auto/Tn) centrado en X en el viewport; stack al fondo.
- Scroll si no cabe a escala mínima (ancho) o si hay más de 8 muros (alto).
- Espesor = color de relleno + etiqueta ``e=``.
- Marca de nivel V2 (disco + cuartos teal) a la izquierda, anclada al pie del fuste.
- Bandas Tn + pie Auto/Tramo/Cont. duales por extremo:
  ``Tn Ini | Auto Ini | elevación | Auto Fin | Tn Fin``.
  Extremo activo a opacidad plena; el otro, tenue (~0.42).
- Fase Mallas: etiqueta por fuste ``M.H.A. e=…`` + ``D.M. / V.=ø… / H.=ø…``
  vía ``mesh_by_wall`` + ``show_mesh_labels``.
"""

from __future__ import division, print_function

ELEV_COL_W_DEFAULT = 300.0
ELEV_COL_W_MIN = 260.0
GUTTER_W = 78.0
CTRL_COL_W = 90.0
BAND_W = 44.0
BAND_GAP = 6.0
LEFT_PAD = 6.0
RIGHT_PAD = 6.0
BTN_W = 56.0
BTN_H = 20.0
PAD = 8.0
# Opacidad de la banda Tn del extremo no configurado (fase activa / último cabezal).
BAND_DIM_OPACITY = 0.42
_FT_TO_MM = 304.8
# Escala mínima legible (~25 px por metro); por debajo → scroll.
_SCALE_PX_PER_MM_MIN = 0.025
# Fit-to-height solo para hasta N muros; con más → escala fija y scrollbar
# (en viewport se ven ~N fustes; el resto queda fuera y se scrollea).
_MAX_WALLS_FIT_SCALE = 8


def _wall_length_mm(m):
    try:
        ft = float(m.get(u"length_ft") or m.get(u"length_u") or 0.0)
    except Exception:
        ft = 0.0
    if ft <= 1e-9:
        ft = 1.0 / _FT_TO_MM
    return max(ft * _FT_TO_MM, 1.0)


def _wall_height_mm(m):
    try:
        h = float(m.get(u"height_mm") or 0.0)
    except Exception:
        h = 0.0
    return max(h, 1.0)


def _wall_fund_height_mm(m):
    try:
        h = float(m.get(u"fund_height_mm") or 0.0)
    except Exception:
        h = 0.0
    if h <= 0.1:
        try:
            fi = m.get(u"fund_info") or {}
            h = float(fi.get(u"height_mm") or 0.0)
        except Exception:
            h = 0.0
    return max(h, 0.0)


def _wall_fund_width_mm(m):
    try:
        w = float(m.get(u"fund_width_mm") or 0.0)
    except Exception:
        w = 0.0
    if w <= 0.1:
        try:
            fi = m.get(u"fund_info") or {}
            w = float(fi.get(u"width_mm") or 0.0)
        except Exception:
            w = 0.0
    return max(w, 0.0)


def _wall_row_height_mm(m):
    return _wall_height_mm(m) + _wall_fund_height_mm(m)


def _compute_elev_layout(meta, viewport_w, viewport_h, legend_h, g_min, span_u):
    """Escala uniforme px/mm; bloque elevación+Auto/Tn centrado en el viewport.

    Orden horizontal (espejo V2 por extremo):
    banda Tn Inicio → pie Auto Ini → elevación → pie Auto Fin → banda Tn Término.
    Escala = fit altura (≤``_MAX_WALLS_FIT_SCALE`` muros) y ancho usable del
    viewport; no se reescala al centrar.
    Tras fijar la escala, la columna de elevación se ajusta al ancho dibujado
    (fustes + gutter de niveles) y el bloque completo se traslada en X para
    quedar centrado en ``vw`` (sobrante negro repartido a ambos lados).
    Si el viewport es más alto que el contenido, el stack se alinea al fondo.
    """
    items = list(meta or [])
    n = len(items)
    if n <= 0:
        n = 1
        items = []
    chrome_h = PAD * 2.0 + float(legend_h or 0.0) + 8.0
    try:
        vw = float(viewport_w or 0.0)
    except Exception:
        vw = 0.0
    try:
        vh = float(viewport_h or 0.0)
    except Exception:
        vh = 0.0

    # Tn+Auto a ambos lados del stack (Inicio izq. / Término der.).
    left_cols_w = BAND_W + BAND_GAP + CTRL_COL_W + BAND_GAP
    right_cols_w = CTRL_COL_W + BAND_GAP + BAND_W + PAD
    # Ancho provisional a pantalla completa solo para calcular escala (fit).
    elev_x_fit = PAD + left_cols_w
    if vw > 40.0:
        elev_col_w_fit = max(
            ELEV_COL_W_MIN, vw - elev_x_fit - right_cols_w,
        )
    else:
        elev_col_w_fit = ELEV_COL_W_DEFAULT

    zone_left = LEFT_PAD + GUTTER_W
    usable_w_fit = max(20.0, elev_col_w_fit - zone_left - RIGHT_PAD)
    avail_h = max(80.0, vh - chrome_h) if vh > 40.0 else 480.0

    lengths_mm = [_wall_length_mm(m) for m in items] if items else [1500.0]
    heights_mm = [_wall_row_height_mm(m) for m in items] if items else [3000.0]
    max_len_mm = max(max(lengths_mm), 1.0)
    # Huella U del stack (pies → mm): incluye offsets entre muros apilados.
    try:
        span_u_f = max(float(span_u or 0.0), 1e-6)
    except Exception:
        span_u_f = 1e-6
    span_mm = max(span_u_f * _FT_TO_MM, max_len_mm)

    # Fit altura: como máximo 8 muros. Con N>8 la escala no sigue bajando
    # (viewport ~8 fustes; el resto → scrollbar). Incluye fundación unida.
    fit_count = min(len(heights_mm), _MAX_WALLS_FIT_SCALE)
    fit_h_mm = max(sum(heights_mm[:fit_count]), 1.0)

    # Escala uniforme: fit (≤8 muros) en altura y huella completa en ancho.
    scale_h = avail_h / fit_h_mm
    scale_w = usable_w_fit / span_mm
    scale = min(scale_h, scale_w)
    if scale < _SCALE_PX_PER_MM_MIN:
        scale = _SCALE_PX_PER_MM_MIN

    row_heights = [max(8.0, h * scale) for h in heights_mm]
    while len(row_heights) < n:
        row_heights.append(max(8.0, 1000.0 * scale))
    stack_h = sum(row_heights[:n])
    content_h = stack_h + chrome_h

    # Columna justa al dibujo a esta escala (sin estirar → hueco negro lateral).
    max_draw_w = max(8.0, span_mm * scale)
    elev_col_w = zone_left + max_draw_w + RIGHT_PAD
    usable_w = max_draw_w
    block_w = left_cols_w + elev_col_w + right_cols_w

    # Centrar el bloque (bandas + elevación + Auto) en el viewport; solo X.
    if vw > 40.0 and block_w < vw:
        block_x = (vw - block_w) / 2.0
        canvas_w = vw
    else:
        block_x = PAD
        canvas_w = block_x + block_w

    band_x_left = block_x
    ctrl_x_left = band_x_left + BAND_W + BAND_GAP
    elev_x = ctrl_x_left + CTRL_COL_W + BAND_GAP
    ctrl_x_right = elev_x + elev_col_w + 8.0
    band_x_right = elev_x + elev_col_w + CTRL_COL_W + BAND_GAP
    canvas_h = max(content_h, vh) if vh > content_h else content_h
    y0 = PAD + float(legend_h or 0.0)
    # Alinear stack al fondo: el sobrante vertical va arriba del stack
    # (pie del último muro cerca del borde inferior − PAD). Con overflow
    # (scroll) no hay slack → y0 queda en el origen natural.
    if canvas_h > content_h:
        y0 += canvas_h - content_h
    return {
        u"elev_x": elev_x,
        u"elev_col_w": elev_col_w,
        u"band_x": band_x_right,  # alias: banda Término (compat)
        u"band_x_left": band_x_left,
        u"band_x_right": band_x_right,
        u"ctrl_x_left": ctrl_x_left,
        u"ctrl_x_right": ctrl_x_right,
        u"row_heights": row_heights,
        u"scale_px_per_mm": scale,
        u"stack_h": stack_h,
        u"canvas_w": canvas_w,
        u"canvas_h": canvas_h,
        u"y0": y0,
        u"content_h": content_h,
        u"g_min": float(g_min or 0.0),
        u"span_u": float(span_u or 1.0),
        u"usable_w": usable_w,
        u"max_len_mm": max_len_mm,
    }

_THICKNESS_UI_PALETTE = (
    u"#4a7a88",
    u"#5a8268",
    u"#7a7348",
    u"#5a6690",
    u"#886070",
    u"#6a5888",
)


def _hex_brush(hex_str, alpha=255):
    from System.Windows.Media import Color, SolidColorBrush

    h = (hex_str or "#000000").lstrip("#")
    aa = max(0, min(255, int(alpha)))
    return SolidColorBrush(
        Color.FromArgb(
            aa,
            int(h[0:2], 16),
            int(h[2:4], 16),
            int(h[4:6], 16),
        )
    )


def _level_theme_brushes():
    """Misma paleta de nivel que Armado Muros V2 ``_level_theme_brushes``."""
    return {
        u"line": _hex_brush("#475569"),
        u"text": _hex_brush("#94a3b8"),
        u"disk": _hex_brush("#cbd5e1"),
        u"bubble": _hex_brush("#22d3ee"),
    }


def _add_level_head_symbol(canv, cx, cy, r=8.0):
    """Símbolo de nivel estilo Revit / Muros V2: disco + cuartos teal opuestos + aro.

    Copia de ``armado_muros_preview_ui._add_level_head_symbol`` (geometría Path),
    sin la línea discontinua (en Machones el leader va a la derecha del símbolo).
    """
    from System.Windows.Controls import Canvas as _Cn
    from System.Windows.Shapes import Ellipse, Path
    from System.Windows.Media import Brushes, Geometry

    br = _level_theme_brushes()
    cx = float(cx)
    cy = float(cy)
    r = float(r)
    bx = cx - r
    by = cy - r
    dia = 2.0 * r

    disk = Ellipse()
    disk.Width = dia
    disk.Height = dia
    disk.Fill = br[u"disk"]
    disk.Stroke = Brushes.Transparent
    _Cn.SetLeft(disk, bx)
    _Cn.SetTop(disk, by)
    canv.Children.Add(disk)

    try:
        p_tl = Path()
        p_tl.Data = Geometry.Parse(
            u"M {0},{1} L {2},{1} A {3},{3} 0 0 1 {0},{4} Z".format(
                cx, cy, cx - r, r, cy - r
            )
        )
        p_tl.Fill = br[u"bubble"]
        canv.Children.Add(p_tl)

        p_br = Path()
        p_br.Data = Geometry.Parse(
            u"M {0},{1} L {2},{1} A {3},{3} 0 0 1 {0},{4} Z".format(
                cx, cy, cx + r, r, cy + r
            )
        )
        p_br.Fill = br[u"bubble"]
        canv.Children.Add(p_br)

        rim = Ellipse()
        rim.Width = dia
        rim.Height = dia
        rim.Fill = Brushes.Transparent
        rim.Stroke = br[u"text"]
        rim.StrokeThickness = 0.95
        _Cn.SetLeft(rim, bx)
        _Cn.SetTop(rim, by)
        canv.Children.Add(rim)
    except Exception:
        # Respaldo mínimo si Path/Geometry no está disponible.
        rim = Ellipse()
        rim.Width = dia
        rim.Height = dia
        rim.Fill = Brushes.Transparent
        rim.Stroke = br[u"bubble"]
        rim.StrokeThickness = 1.0
        _Cn.SetLeft(rim, bx)
        _Cn.SetTop(rim, by)
        canv.Children.Add(rim)

    return cx, r


def _level_label_text(m):
    """Texto numérico de cota de nivel (metros, ref. activa en meta) para la marca izquierda."""
    lvl_label = m.get(u"level_z_m") if m is not None else None
    if lvl_label is None:
        try:
            return u"{0:.3f}".format(float((m or {}).get(u"z_mm") or 0.0) / 1000.0)
        except Exception:
            return (m or {}).get(u"label") or u""
    try:
        return u"{0:.3f}".format(float(lvl_label))
    except Exception:
        return u"{0}".format(lvl_label)


def _draw_level_marker(canvas, elev_x, elev_col_w, foot_y, label_txt):
    """Marca de nivel (texto + disco + guía) a la izquierda del fuste.

    Debe dibujarse *después* de todos los ``bg`` de fila: el símbolo se ancla
    en ``foot_y`` (borde inferior del fuste) y la mitad inferior queda en la
    fila siguiente; si el ``bg`` opaco de esa fila se pinta encima, se recorta.
    """
    from System.Windows.Controls import Canvas, TextBlock
    from System.Windows.Shapes import Line
    from System.Windows import FontWeights, TextAlignment

    br_lvl = _level_theme_brushes()
    lvl_txt = TextBlock()
    lvl_txt.Text = u"{0}".format(label_txt)
    lvl_txt.Foreground = br_lvl[u"text"]
    lvl_txt.FontSize = 9.0
    try:
        lvl_txt.FontWeight = FontWeights.SemiBold
    except Exception:
        pass
    try:
        lvl_txt.TextAlignment = TextAlignment.Right
    except Exception:
        pass
    lvl_txt_w = 32.0
    lvl_left = float(elev_x) + 4.0
    Canvas.SetLeft(lvl_txt, lvl_left)
    Canvas.SetTop(lvl_txt, float(foot_y) - 7.0)
    try:
        lvl_txt.Width = lvl_txt_w
    except Exception:
        pass
    canvas.Children.Add(lvl_txt)

    bubble_r = 8.0
    bubble_cx = lvl_left + lvl_txt_w + 4.0 + bubble_r
    _add_level_head_symbol(canvas, bubble_cx, float(foot_y), bubble_r)

    lead = Line()
    lead.X1 = bubble_cx + bubble_r + 2.0
    lead.X2 = float(elev_x) + float(elev_col_w)
    lead.Y1 = float(foot_y)
    lead.Y2 = float(foot_y)
    lead.Stroke = br_lvl[u"line"]
    lead.StrokeThickness = 0.9
    try:
        from System.Windows.Media import DoubleCollection

        dc = DoubleCollection()
        dc.Add(4.0)
        dc.Add(3.0)
        lead.StrokeDashArray = dc
    except Exception:
        pass
    if lead.X2 > lead.X1 + 4.0:
        canvas.Children.Add(lead)


def lap_mm_from_long_diam(diam_mm):
    d = float(diam_mm or 16.0)
    if d <= 10:
        return 350.0
    if d <= 12:
        return 420.0
    if d <= 16:
        return 480.0
    if d <= 20:
        return 560.0
    return 700.0


def _type_label_clean(raw):
    try:
        s = unicode(raw or u"").strip()
    except Exception:
        s = u"{0}".format(raw or u"").strip()
    if not s:
        return u"Muro"
    low = s.lower()
    idx = low.find(u" e=")
    if idx >= 0:
        return s[:idx].strip() or u"Muro"
    return s


def _mesh_bar_spec_txt(diam_mm, spacing_mm):
    """Texto de barra malla: ``ø10@200``."""
    try:
        d = int(round(float(diam_mm or 10)))
    except Exception:
        d = 10
    try:
        s = int(round(float(spacing_mm or 200)))
    except Exception:
        s = 200
    return u"ø{0}@{1}".format(d, s)


def _fill_wall_fuste_label(
    sp,
    m,
    thick_key,
    conf,
    mesh_cfg,
    show_mesh_labels,
    br_title,
    br_wall,
    br_muted,
    br_accent_mesh,
    hx,
    TextBlock,
    StackPanel,
    FontWeights,
    HorizontalAlignment,
    VerticalAlignment,
    TextAlignment,
    Thickness,
    Orientation,
):
    """Rellena el StackPanel del fuste: etiqueta malla o confinamiento legacy."""
    tipo = _type_label_clean(m.get(u"type_name") or m.get(u"name") or u"Muro")
    try:
        while sp.Children.Count > 0:
            sp.Children.RemoveAt(0)
    except Exception:
        pass

    if show_mesh_labels:
        # Línea superior: «M.H.A. e=300» en accent (#5BC0DE).
        t_top = TextBlock()
        t_top.Text = u"{0}  e={1}".format(tipo, thick_key)
        t_top.Foreground = br_accent_mesh
        t_top.FontSize = 10.0
        t_top.FontWeight = FontWeights.Bold
        t_top.TextAlignment = TextAlignment.Center
        t_top.HorizontalAlignment = HorizontalAlignment.Center
        sp.Children.Add(t_top)

        mesh = mesh_cfg or {}
        activo = True
        try:
            if u"activo" in mesh:
                activo = bool(mesh.get(u"activo"))
        except Exception:
            activo = True

        if not activo:
            t_off = TextBlock()
            t_off.Text = u"sin malla"
            t_off.Foreground = br_muted
            t_off.FontSize = 9.0
            t_off.FontWeight = FontWeights.SemiBold
            t_off.TextAlignment = TextAlignment.Center
            t_off.HorizontalAlignment = HorizontalAlignment.Center
            t_off.Opacity = 0.85
            sp.Children.Add(t_off)
            return tipo

        # Bloque D.M. | V./H. (doble malla ext+int).
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.HorizontalAlignment = HorizontalAlignment.Center
        row.Margin = Thickness(0, 2, 0, 0)

        doble = True
        try:
            if u"doble" in mesh:
                doble = bool(mesh.get(u"doble"))
        except Exception:
            doble = True
        if doble:
            t_dm = TextBlock()
            t_dm.Text = u"D.M."
            t_dm.Foreground = br_title
            t_dm.FontSize = 9.0
            t_dm.FontWeight = FontWeights.Bold
            t_dm.VerticalAlignment = VerticalAlignment.Center
            t_dm.Margin = Thickness(0, 0, 6, 0)
            row.Children.Add(t_dm)

        vh = StackPanel()
        vh.VerticalAlignment = VerticalAlignment.Center
        t_v = TextBlock()
        t_v.Text = u"V.={0}".format(
            _mesh_bar_spec_txt(
                mesh.get(u"v_diam_mm"), mesh.get(u"v_spacing_mm"),
            )
        )
        t_v.Foreground = br_title
        t_v.FontSize = 9.0
        t_v.FontWeight = FontWeights.SemiBold
        t_v.TextAlignment = TextAlignment.Left
        t_h = TextBlock()
        t_h.Text = u"H.={0}".format(
            _mesh_bar_spec_txt(
                mesh.get(u"h_diam_mm"), mesh.get(u"h_spacing_mm"),
            )
        )
        t_h.Foreground = br_title
        t_h.FontSize = 9.0
        t_h.FontWeight = FontWeights.SemiBold
        t_h.TextAlignment = TextAlignment.Left
        vh.Children.Add(t_v)
        vh.Children.Add(t_h)
        row.Children.Add(vh)
        sp.Children.Add(row)
        return tipo

    # Legacy: tipo / e= / Ø@ confinamiento.
    t1 = TextBlock()
    t1.Text = tipo
    t1.Foreground = br_title
    t1.FontSize = 10.0
    t1.FontWeight = FontWeights.SemiBold
    t1.TextAlignment = TextAlignment.Center
    t1.HorizontalAlignment = HorizontalAlignment.Center
    t2 = TextBlock()
    t2.Text = u"e={0}".format(thick_key)
    t2.Foreground = _hex_brush(hx, 230)
    t2.FontSize = 10.0
    t2.FontWeight = FontWeights.SemiBold
    t2.TextAlignment = TextAlignment.Center
    t2.HorizontalAlignment = HorizontalAlignment.Center
    t3 = TextBlock()
    try:
        cd = int(round(float(conf.get(u"diam_mm") or 10)))
        cs = int(round(float(conf.get(u"spacing_mm") or 200)))
        t3.Text = u"Ø{0}@{1}".format(cd, cs)
    except Exception:
        t3.Text = u"{0}".format(m.get(u"name") or u"Muro")
    t3.Foreground = br_wall
    t3.FontSize = 9.0
    t3.TextAlignment = TextAlignment.Center
    t3.HorizontalAlignment = HorizontalAlignment.Center
    sp.Children.Add(t1)
    sp.Children.Add(t2)
    sp.Children.Add(t3)
    return tipo


def _thickness_color_map(wall_meta):
    thicknesses = set()
    for m in wall_meta or []:
        try:
            thicknesses.add(int(round(float(m.get(u"thick_mm") or 200))))
        except Exception:
            thicknesses.add(200)
    if not thicknesses:
        thicknesses.add(200)
    mp = {}
    for i, th in enumerate(sorted(thicknesses)):
        mp[int(th)] = _THICKNESS_UI_PALETTE[i % len(_THICKNESS_UI_PALETTE)]
    return mp


def _u_layout_metrics(wall_meta):
    """Escala U compartida: span = huella del stack (máx. largo real)."""
    items = list(wall_meta or [])
    if not items:
        return 0.0, 1.0, 1.0
    u_mins = []
    u_maxs = []
    max_len = 1e-6
    for m in items:
        try:
            u0 = float(m.get(u"u_start", 0.0))
            lu = max(float(m.get(u"length_u", m.get(u"length_ft", 1.0))), 1e-6)
        except Exception:
            u0, lu = 0.0, 1.0
        u_mins.append(u0)
        u_maxs.append(u0 + lu)
        if lu > max_len:
            max_len = lu
    g_min = min(u_mins)
    g_max = max(u_maxs)
    span = max(g_max - g_min, max_len, 1e-6)
    return g_min, g_max, span


def _normalize_extremo_key(extremo, default=u"inicio"):
    """``inicio`` / ``fin`` canónicos (LocationCurve extremo, no lado canvas)."""
    act = (extremo or default).strip().lower()
    if act in (u"fin", u"termino", u"término", u"final"):
        return u"fin"
    return u"inicio"


def _extremo_side_label(extremo_key):
    return u"Inicio" if extremo_key == u"inicio" else u"Término"


def _draw_w_xoff(m, elev_w, scale_px_per_mm, g_min=None, span_u=None):
    """Ancho/offset del prisma según posición U relativa (como V2 stacked).

    ``draw_w = length_mm * scale``; ``x_off`` desde ``(u_start - g_min)``
    en la misma escala, para respetar offsets entre muros apilados.
    Si falta huella U válida, centra (fallback).
    """
    zone_left = LEFT_PAD + GUTTER_W
    zone_w = max(20.0, float(elev_w) - zone_left - RIGHT_PAD)
    length_mm = _wall_length_mm(m)
    scale = max(float(scale_px_per_mm or 0.0), 1e-9)
    draw_w = max(8.0, length_mm * scale)

    use_u = False
    u0 = 0.0
    gmin = 0.0
    try:
        if m is not None and m.get(u"u_start") is not None:
            u0 = float(m.get(u"u_start"))
            gmin = float(g_min if g_min is not None else 0.0)
            if span_u is not None and float(span_u) > 1e-9:
                use_u = True
    except Exception:
        use_u = False

    if use_u:
        x_off = zone_left + (u0 - gmin) * _FT_TO_MM * scale
        wall_zone_max = zone_left + zone_w
        if x_off < zone_left:
            x_off = zone_left
        if x_off + draw_w > wall_zone_max + 0.5:
            x_off = max(zone_left, wall_zone_max - draw_w)
    else:
        x_off = zone_left + (zone_w - draw_w) / 2.0
    return x_off, draw_w


def _footing_elev_width_px(draw_w, scale_px, fund_width_mm):
    """Ancho dibujado de zapata: muro + voladizo visible, no bbox completo de losa.

    Fórmula:
    - ``wall_px = draw_w`` (largo elevación del muro).
    - Voladizo por lado ≈ 13 % de ``wall_px`` (mín. 3 px), tope 15 % por lado
      → ``default = wall + 2·margin`` (~1.26×), ``max_foot = wall · 1.30``.
    - Si hay ``fund_width_mm`` (bbox) y es modesto (≤ ``max_foot · 1.25``):
      ``foot_w = max(wall, min(bbox_px, max_foot))``.
    - Si bbox falta o es enorme (Foundation Slab multi-muro): ``default``.
    - Altura sigue fuera: ``height_mm * scale``; centrado bajo el fuste en el caller.
    """
    wall_px = max(8.0, float(draw_w))
    overhang_frac = 0.13
    max_overhang_frac = 0.15
    margin_each = max(wall_px * overhang_frac, 3.0)
    margin_each = min(margin_each, wall_px * max_overhang_frac)
    default_foot = wall_px + 2.0 * margin_each
    max_foot = wall_px * (1.0 + 2.0 * max_overhang_frac)

    try:
        fw_mm = float(fund_width_mm or 0.0)
    except Exception:
        fw_mm = 0.0
    if fw_mm > 0.1:
        try:
            bbox_px = fw_mm * max(float(scale_px or 0.0), 0.0)
        except Exception:
            bbox_px = 0.0
        if bbox_px > 0.5 and bbox_px <= max_foot * 1.25:
            return max(8.0, max(wall_px, min(bbox_px, max_foot)))
    return max(8.0, default_foot)


def _draw_foundation_footing_elev(
    canvas,
    elev_x,
    x_off,
    draw_w,
    y_stem_foot,
    scale_px,
    fund_height_mm,
    fund_width_mm,
    color_hex,
    fund_info,
    Canvas,
    TextBlock,
    FontWeights,
):
    """Zapata/losa bajo el pie del fuste; alto a escala real, largo recortado al muro."""
    from System.Windows.Shapes import Rectangle as WRect, Line as WLine
    from System.Windows import HorizontalAlignment

    try:
        h_mm = float(fund_height_mm or 0.0)
    except Exception:
        h_mm = 0.0
    if h_mm <= 0.1 and fund_info:
        try:
            h_mm = float(fund_info.get(u"height_mm") or 0.0)
        except Exception:
            h_mm = 0.0
    try:
        fund_h_px = max(2.0, float(h_mm) * float(scale_px or 0.0)) if h_mm > 0.1 else 0.0
    except Exception:
        fund_h_px = 0.0
    if fund_h_px <= 0.5:
        return 0.0

    hx = u"#64748b"  # gris — no reutilizar color de espesor (parecía muro fantasma)
    fill_br = _hex_brush(hx, 110)
    stroke_br = _hex_brush(hx, 220)
    text_br = _hex_brush(u"#94a3b8", 220)

    foot_w = _footing_elev_width_px(draw_w, scale_px, fund_width_mm)
    foot_x = float(elev_x) + float(x_off) - (foot_w - float(draw_w)) * 0.5
    foot_y = float(y_stem_foot)

    foot = WRect()
    foot.Width = foot_w
    foot.Height = fund_h_px
    foot.Fill = fill_br
    foot.Stroke = stroke_br
    foot.StrokeThickness = 1.0
    Canvas.SetLeft(foot, foot_x)
    Canvas.SetTop(foot, foot_y)
    canvas.Children.Add(foot)

    joint = WLine()
    joint.X1 = foot_x + 1.0
    joint.X2 = foot_x + foot_w - 1.0
    joint.Y1 = foot_y
    joint.Y2 = foot_y
    joint.Stroke = stroke_br
    joint.StrokeThickness = 1.2
    canvas.Children.Add(joint)

    cap_h = h_mm
    if (cap_h is None or cap_h <= 0.1) and fund_info:
        try:
            cap_h = float(fund_info.get(u"height_mm") or fund_height_mm or 0.0)
        except Exception:
            cap_h = None
    if cap_h is not None and cap_h > 0.1:
        cap = u"Fund. H\u2248{:.0f}".format(cap_h)
    else:
        n_f = int((fund_info or {}).get(u"count", 1) or 1)
        cap = u"Fund." if n_f <= 1 else u"Fund. \u00d7{}".format(n_f)

    tb = TextBlock()
    tb.Text = cap
    tb.Foreground = text_br
    tb.FontSize = 8.0
    tb.FontWeight = FontWeights.SemiBold
    tb.HorizontalAlignment = HorizontalAlignment.Center
    Canvas.SetLeft(tb, foot_x)
    Canvas.SetTop(tb, foot_y + max(1.0, (fund_h_px - 10.0) * 0.5))
    tb.Width = foot_w
    canvas.Children.Add(tb)
    return fund_h_px


def _build_tramo_band_blocks(display, segments):
    """Bloques Tn en orden visual (arriba = Z alto), fusionando filas del mismo tramo."""

    def _seg_id(wi):
        for s in segments or []:
            if wi in (s.get(u"wall_indices") or []):
                try:
                    return int(s.get(u"id", -1))
                except Exception:
                    return -1
        return -1

    blocks = []
    i = 0
    while i < len(display):
        sid = _seg_id(display[i])
        j = i + 1
        while j < len(display):
            if _seg_id(display[j]) != sid:
                break
            j += 1
        blocks.append({u"sid": sid, u"di0": i, u"count": j - i})
        i = j
    return blocks


def _selected_segment_id_set(selected_segment, selected_segments=None):
    """Conjunto de ids de tramo seleccionados (multi) + primario de respaldo."""
    out = set()
    for x in list(selected_segments or []):
        try:
            out.add(int(x))
        except Exception:
            pass
    try:
        out.add(int(selected_segment))
    except Exception:
        pass
    return out


def _draw_tramo_band_column(
    canvas,
    band_x,
    display,
    y_at_di,
    row_heights,
    y0,
    segments,
    selected_segment,
    cfgs,
    focus,
    band_active,
    dim_opacity,
    on_select_segment,
    extremo_label,
    br_elev,
    br_border,
    br_sel,
    br_title,
    br_muted,
    Border,
    StackPanel,
    TextBlock,
    ToolTip,
    Thickness,
    FontWeights,
    HorizontalAlignment,
    VerticalAlignment,
    TextAlignment,
    MouseButtonEventHandler,
    Canvas,
    selected_segments=None,
):
    """Dibuja una columna de bandas Tn en ``band_x`` (Inicio o Término)."""
    try:
        sel_seg = int(selected_segment)
    except Exception:
        sel_seg = -1
    sel_set = _selected_segment_id_set(sel_seg, selected_segments)
    try:
        opac = float(dim_opacity if not band_active else 1.0)
    except Exception:
        opac = 1.0 if band_active else BAND_DIM_OPACITY
    opac = max(0.15, min(1.0, opac))
    label = extremo_label or u"Tramo"

    for bi, blk in enumerate(_build_tramo_band_blocks(display, segments)):
        sid = int(blk.get(u"sid", -1))
        count = int(blk.get(u"count", 1))
        di0 = int(blk.get(u"di0", 0))
        y_band = y_at_di[di0] if di0 < len(y_at_di) else float(y0)
        h_band = 0.0
        for _k in range(di0, min(di0 + count, len(display))):
            _wi_b = display[_k]
            h_band += float(row_heights[_wi_b] if _wi_b < len(row_heights) else 40.0)
        block_sel = (
            sid in sel_set and focus == u"tramo" and band_active
        )
        is_primary = block_sel and sid == sel_seg
        band = Border()
        band.Width = BAND_W
        band.Height = h_band
        if is_primary:
            band.Background = _hex_brush("#5BC0DE", 56)
            band.BorderBrush = br_sel
        elif block_sel:
            band.Background = _hex_brush("#5BC0DE", 32)
            band.BorderBrush = br_sel
        else:
            band.Background = br_elev
            band.BorderBrush = br_border
        if bi == 0:
            band.BorderThickness = Thickness(1.5 if is_primary else 1.0)
        else:
            band.BorderThickness = Thickness(
                1.5 if is_primary else 1.0,
                0.0,
                1.5 if is_primary else 1.0,
                1.5 if is_primary else 1.0,
            )
        try:
            band.Opacity = opac
        except Exception:
            pass
        tip_b = ToolTip()
        if sid >= 0:
            tip_b.Content = (
                u"{0} · Tramo T{1}: clic = seleccionar · "
                u"Ctrl+clic = multi · Mayús+clic = rango"
            ).format(label, sid + 1)
        else:
            tip_b.Content = u"{0} · Sin tramo".format(label)
        band.ToolTip = tip_b
        Canvas.SetLeft(band, band_x)
        Canvas.SetTop(band, y_band)

        sp_b = StackPanel()
        sp_b.HorizontalAlignment = HorizontalAlignment.Center
        sp_b.VerticalAlignment = VerticalAlignment.Center
        t_tn = TextBlock()
        t_tn.Text = u"T{0}".format(sid + 1) if sid >= 0 else u"—"
        t_tn.Foreground = br_sel if block_sel else br_title
        t_tn.FontSize = 12.0
        t_tn.FontWeight = FontWeights.Bold if is_primary else FontWeights.SemiBold
        t_tn.TextAlignment = TextAlignment.Center
        t_tn.HorizontalAlignment = HorizontalAlignment.Center
        if count >= 2:
            try:
                from System.Windows.Media import RotateTransform

                t_tn.LayoutTransform = RotateTransform(-90.0)
            except Exception:
                pass
        sp_b.Children.Add(t_tn)
        if 0 <= sid < len(cfgs):
            sc = cfgs[sid] or {}
            t_ab = TextBlock()
            t_ab.Text = u"{0}×{1}".format(
                int(sc.get(u"bars_a") or 4),
                int(sc.get(u"bars_b") or 6),
            )
            t_ab.Foreground = br_muted
            t_ab.FontSize = 8.0
            t_ab.TextAlignment = TextAlignment.Center
            t_ab.HorizontalAlignment = HorizontalAlignment.Center
            if count >= 2:
                try:
                    from System.Windows.Media import RotateTransform

                    t_ab.LayoutTransform = RotateTransform(-90.0)
                except Exception:
                    pass
            sp_b.Children.Add(t_ab)
        band.Child = sp_b

        if sid >= 0 and on_select_segment is not None:
            try:

                def _mk_band(sid_):
                    def _h(sender, args):
                        on_select_segment(int(sid_))

                    return _h

                band.MouseLeftButtonDown += MouseButtonEventHandler(_mk_band(sid))
            except Exception:
                pass
        canvas.Children.Add(band)


def redraw_elevation(
    canvas,
    wall_meta,
    troceo_modes,
    segments,
    selected_segment,
    long_diam_mm,
    on_cycle_wall,
    on_select_segment,
    selected_wall=0,
    selected_walls=None,
    rail_focus=u"tramo",
    conf_by_wall=None,
    on_select_wall=None,
    cfg_by_segment=None,
    viewport_w=None,
    viewport_h=None,
    on_draw_extremo_marks=None,
    segments_inicio=None,
    selected_segment_inicio=None,
    cfg_by_segment_inicio=None,
    on_select_segment_inicio=None,
    segments_fin=None,
    selected_segment_fin=None,
    cfg_by_segment_fin=None,
    on_select_segment_fin=None,
    troceo_modes_inicio=None,
    troceo_modes_fin=None,
    on_cycle_wall_inicio=None,
    on_cycle_wall_fin=None,
    active_extremo=u"inicio",
    extremo_left=None,
    extremo_right=None,
    band_dim_opacity=None,
    mesh_by_wall=None,
    show_mesh_labels=False,
    selected_segments=None,
    selected_segments_inicio=None,
    selected_segments_fin=None,
):
    """
    ``wall_meta``: lista ordenada base→cima con keys
    height_mm, thick_mm, label, name, type_name, u_start, length_u.

    Dual rail V3 (estado por extremo como V2):
    ``Tn Ini | Auto Ini | elevación | Auto Fin | Tn Fin``.
    ``troceo_modes_inicio`` / ``troceo_modes_fin`` independientes;
    ``active_extremo`` → opacidad plena en ese lado (banda + pie Auto).
    ``extremo_left`` / ``extremo_right`` mapean qué extremo (inicio/fin)
    va en cada banda física; deben alinear +U del canvas con
    ``RightDirection`` de la vista (vía ``cabezal_extremos_en_lados_*``).
    Empalme visual / highlight de fuste siguen el extremo activo.
    Si no se pasan ``segments_inicio``/``segments_fin``, cae al modo
    legacy de una sola banda/pie a la derecha.

    ``mesh_by_wall`` (fase Mallas): lista paralela a ``wall_meta`` con keys
    ``activo``, ``doble``, ``v_diam_mm``, ``v_spacing_mm``, ``h_diam_mm``,
    ``h_spacing_mm``. Con ``show_mesh_labels=True`` el fuste muestra
    ``M.H.A. e=…`` + bloque ``D.M. / V.=ø… / H.=ø…`` (o «sin malla»).
    """
    if canvas is None:
        return
    try:
        from System.Windows.Controls import (
            Canvas,
            Button,
            TextBlock,
            Border,
            ToolTip,
            StackPanel,
            Orientation,
        )
        from System.Windows.Shapes import Rectangle, Line
        from System.Windows import (
            Thickness,
            FontWeights,
            CornerRadius,
            HorizontalAlignment,
            VerticalAlignment,
            TextAlignment,
        )
        from System.Windows.Input import MouseButtonEventHandler
        from System.Windows import RoutedEventHandler
    except Exception:
        return

    try:
        canvas.Children.Clear()
    except Exception:
        pass

    meta = list(wall_meta or [])
    n = len(meta)
    if n <= 0:
        tb = TextBlock()
        tb.Text = u"No hay muros en el stack."
        tb.Foreground = _hex_brush("#64748b")
        Canvas.SetLeft(tb, 12)
        Canvas.SetTop(tb, 12)
        canvas.Children.Add(tb)
        canvas.Height = 80
        return

    from armado_muros_v3_troceo import (
        TROCEO_AUTO,
        effective_empalme,
        empalme_indices_from_modes,
        pie_caption,
        compute_auto_troceo_flags_from_meta,
        lowest_z_meta_index,
        log_elev_troceo_debug,
    )

    def _pad_modes(src):
        out = list(src or [])
        while len(out) < n:
            out.append(TROCEO_AUTO)
        if len(out) > n:
            out = out[:n]
        return out

    modes_legacy = _pad_modes(troceo_modes)
    # Base (Z más baja) fuera; Auto desde el 2.º muro apilado.
    autos = compute_auto_troceo_flags_from_meta(meta)
    base_i = lowest_z_meta_index(meta)
    if base_i is None:
        base_i = 0
    # Nunca empalme / auto visual en el muro min-Z.
    if 0 <= base_i < n:
        autos[base_i] = False

    display = list(reversed(range(n)))
    color_mp = _thickness_color_map(meta)
    g_min, _g_max, span = _u_layout_metrics(meta)
    confs = list(conf_by_wall or [])
    while len(confs) < n:
        confs.append({u"diam_mm": 10.0, u"spacing_mm": 200.0})
    meshes = list(mesh_by_wall or [])
    while len(meshes) < n:
        meshes.append({
            u"activo": True,
            u"doble": True,
            u"v_diam_mm": 10.0,
            u"v_spacing_mm": 200.0,
            u"h_diam_mm": 10.0,
            u"h_spacing_mm": 200.0,
        })
    show_mesh = bool(show_mesh_labels)

    dual_bands = (
        segments_inicio is not None and segments_fin is not None
    )
    if dual_bands:
        segs_ini = list(segments_inicio or [])
        segs_fin = list(segments_fin or [])
        try:
            sel_ini = int(
                selected_segment_inicio
                if selected_segment_inicio is not None
                else (selected_segment or 0)
            )
        except Exception:
            sel_ini = 0
        try:
            sel_fin = int(
                selected_segment_fin
                if selected_segment_fin is not None
                else (selected_segment or 0)
            )
        except Exception:
            sel_fin = 0
        sels_ini = _selected_segment_id_set(sel_ini, selected_segments_inicio)
        sels_fin = _selected_segment_id_set(sel_fin, selected_segments_fin)
        cfgs_ini = list(cfg_by_segment_inicio or [])
        cfgs_fin = list(cfg_by_segment_fin or [])
        act = _normalize_extremo_key(active_extremo, u"inicio")
        ex_left = _normalize_extremo_key(extremo_left, u"inicio")
        ex_right = _normalize_extremo_key(extremo_right, u"fin")
        modes_ini = _pad_modes(
            troceo_modes_inicio
            if troceo_modes_inicio is not None
            else modes_legacy
        )
        modes_fin = _pad_modes(
            troceo_modes_fin
            if troceo_modes_fin is not None
            else modes_legacy
        )
        cycle_ini = (
            on_cycle_wall_inicio
            if on_cycle_wall_inicio is not None
            else on_cycle_wall
        )
        cycle_fin = (
            on_cycle_wall_fin
            if on_cycle_wall_fin is not None
            else on_cycle_wall
        )
        on_sel_ini = (
            on_select_segment_inicio
            if on_select_segment_inicio is not None
            else on_select_segment
        )
        on_sel_fin = (
            on_select_segment_fin
            if on_select_segment_fin is not None
            else on_select_segment
        )
        by_ex = {
            u"inicio": {
                u"segs": segs_ini,
                u"sel": sel_ini,
                u"sels": sels_ini,
                u"cfgs": cfgs_ini,
                u"modes": modes_ini,
                u"cycle": cycle_ini,
                u"on_sel": on_sel_ini,
            },
            u"fin": {
                u"segs": segs_fin,
                u"sel": sel_fin,
                u"sels": sels_fin,
                u"cfgs": cfgs_fin,
                u"modes": modes_fin,
                u"cycle": cycle_fin,
                u"on_sel": on_sel_fin,
            },
        }
        side_left = by_ex[ex_left]
        side_right = by_ex[ex_right]
        # Highlight de fuste / empalme siguen el extremo activo (inicio/fin).
        active_pack = by_ex[act]
        segments = active_pack[u"segs"]
        selected_segment = active_pack[u"sel"]
        selected_set = active_pack[u"sels"]
        cfgs = active_pack[u"cfgs"]
        modes = active_pack[u"modes"]
    else:
        segs_ini = segs_fin = None
        sel_ini = sel_fin = 0
        sels_ini = sels_fin = set()
        cfgs_ini = cfgs_fin = []
        modes_ini = modes_fin = modes_legacy
        act = u"inicio"
        segments = list(segments or [])
        cfgs = list(cfg_by_segment or [])
        modes = modes_legacy
        cycle_ini = cycle_fin = on_cycle_wall
        selected_set = _selected_segment_id_set(selected_segment, selected_segments)

    try:
        log_elev_troceo_debug(meta, autos=autos, modes=modes)
    except Exception:
        pass

    try:
        sel_wall = int(selected_wall or 0)
    except Exception:
        sel_wall = 0
    sel_walls = set()
    for x in list(selected_walls or []):
        try:
            sel_walls.add(int(x))
        except Exception:
            pass
    if not sel_walls:
        sel_walls.add(sel_wall)
    focus = rail_focus or u"tramo"
    try:
        dim_op = float(
            band_dim_opacity
            if band_dim_opacity is not None
            else BAND_DIM_OPACITY
        )
    except Exception:
        dim_op = BAND_DIM_OPACITY

    legend_h = 0.0
    uniq_th = sorted(color_mp.keys())
    if len(uniq_th) > 1:
        legend_h = 18.0

    lay = _compute_elev_layout(meta, viewport_w, viewport_h, legend_h, g_min, span)
    elev_x = float(lay[u"elev_x"])
    elev_col_w = float(lay[u"elev_col_w"])

    def _lay_f(key, default):
        """Lee float de layout; no tratar 0.0 como ausente (``or`` sería falso)."""
        if key in lay and lay[key] is not None:
            try:
                return float(lay[key])
            except Exception:
                pass
        try:
            return float(default)
        except Exception:
            return 0.0

    band_x_left = _lay_f(u"band_x_left", elev_x - BAND_GAP - BAND_W)
    band_x_right = _lay_f(
        u"band_x_right",
        lay[u"band_x"] if lay.get(u"band_x") is not None
        else (elev_x + elev_col_w + CTRL_COL_W + BAND_GAP),
    )
    ctrl_x_left = _lay_f(u"ctrl_x_left", band_x_left + BAND_W + BAND_GAP)
    ctrl_x_right = _lay_f(u"ctrl_x_right", elev_x + elev_col_w + 8.0)
    row_heights = list(lay[u"row_heights"] or [])
    while len(row_heights) < n:
        row_heights.append(40.0)
    scale_px = float(lay[u"scale_px_per_mm"] or 0.05)
    lay_g_min = _lay_f(u"g_min", g_min)
    lay_span_u = _lay_f(u"span_u", span)
    canvas.Width = max(40.0, float(lay[u"canvas_w"] or 40.0))
    canvas.Height = max(40.0, float(lay[u"canvas_h"] or 40.0))

    # Y acumulado por índice de display (cima → base).
    y_at_di = []
    _yc = float(lay[u"y0"])
    for _di, _wi in enumerate(display):
        y_at_di.append(_yc)
        _yc += float(row_heights[_wi] if _wi < len(row_heights) else 40.0)

    br_border = _hex_brush("#21465C")
    br_title = _hex_brush("#E8F4F8")
    br_muted = _hex_brush("#64748b")
    br_accent = _hex_brush("#22d3ee")
    br_level = _hex_brush("#94a3b8")
    br_red = _hex_brush("#ef4444")
    br_a = _hex_brush("#38bdf8")
    br_b = _hex_brush("#f59e0b")
    br_app = _hex_brush("#071018")
    br_manual_bg = _hex_brush("#164e63")
    br_sel = _hex_brush("#5BC0DE")
    br_mesh_accent = br_sel  # etiqueta malla: tipo + e= (#5BC0DE)
    br_wall = _hex_brush("#4ade80")
    br_elev = _hex_brush("#0E1B32")

    y0 = float(lay[u"y0"])

    if legend_h > 0:
        lx = elev_x
        leg_top = max(PAD, y0 - legend_h)
        tb_leg = TextBlock()
        tb_leg.Text = u"Espesor:"
        tb_leg.Foreground = br_muted
        tb_leg.FontSize = 9.0
        tb_leg.FontWeight = FontWeights.SemiBold
        Canvas.SetLeft(tb_leg, lx)
        Canvas.SetTop(tb_leg, leg_top)
        canvas.Children.Add(tb_leg)
        lx += 48.0
        for th in uniq_th:
            hx = color_mp[th]
            sw = Border()
            sw.Width = 10.0
            sw.Height = 10.0
            sw.CornerRadius = CornerRadius(2.0)
            sw.Background = _hex_brush(hx, 220)
            sw.BorderBrush = _hex_brush(hx, 255)
            sw.BorderThickness = Thickness(1)
            Canvas.SetLeft(sw, lx)
            Canvas.SetTop(sw, leg_top + 2.0)
            canvas.Children.Add(sw)
            tl = TextBlock()
            tl.Text = u"e={0}".format(int(th))
            tl.Foreground = br_level
            tl.FontSize = 9.0
            Canvas.SetLeft(tl, lx + 14.0)
            Canvas.SetTop(tl, leg_top)
            canvas.Children.Add(tl)
            lx += 52.0

    def _seg_of(wi):
        for s in segments or []:
            if wi in (s.get(u"wall_indices") or []):
                return s
        return None

    # --- Bandas Tn: dual (Inicio izq + Término der) o legacy (solo der) ---
    _band_kw = dict(
        display=display,
        y_at_di=y_at_di,
        row_heights=row_heights,
        y0=y0,
        focus=focus,
        dim_opacity=dim_op,
        br_elev=br_elev,
        br_border=br_border,
        br_sel=br_sel,
        br_title=br_title,
        br_muted=br_muted,
        Border=Border,
        StackPanel=StackPanel,
        TextBlock=TextBlock,
        ToolTip=ToolTip,
        Thickness=Thickness,
        FontWeights=FontWeights,
        HorizontalAlignment=HorizontalAlignment,
        VerticalAlignment=VerticalAlignment,
        TextAlignment=TextAlignment,
        MouseButtonEventHandler=MouseButtonEventHandler,
        Canvas=Canvas,
    )
    if dual_bands:
        _draw_tramo_band_column(
            canvas,
            band_x_left,
            segments=side_left[u"segs"],
            selected_segment=side_left[u"sel"],
            selected_segments=side_left[u"sels"],
            cfgs=side_left[u"cfgs"],
            band_active=(act == ex_left),
            on_select_segment=side_left[u"on_sel"],
            extremo_label=_extremo_side_label(ex_left),
            **_band_kw
        )
        _draw_tramo_band_column(
            canvas,
            band_x_right,
            segments=side_right[u"segs"],
            selected_segment=side_right[u"sel"],
            selected_segments=side_right[u"sels"],
            cfgs=side_right[u"cfgs"],
            band_active=(act == ex_right),
            on_select_segment=side_right[u"on_sel"],
            extremo_label=_extremo_side_label(ex_right),
            **_band_kw
        )
    else:
        _draw_tramo_band_column(
            canvas,
            band_x_right,
            segments=segments,
            selected_segment=selected_segment,
            selected_segments=selected_set,
            cfgs=cfgs,
            band_active=True,
            on_select_segment=on_select_segment,
            extremo_label=u"Tramo",
            **_band_kw
        )

    y_bottom_of_wall = {}
    band_geom = {}
    # Marcas de nivel al final: foot_y coincide con el top del bg de la fila
    # inferior; si se pintan en el mismo loop quedan tapadas (excepto la base).
    level_marks = []

    for di, wi in enumerate(display):
        m = meta[wi]
        row_h = float(row_heights[wi] if wi < len(row_heights) else 40.0)
        y = y_at_di[di] if di < len(y_at_di) else float(lay[u"y0"])
        seg = _seg_of(wi)
        try:
            seg_id_i = int(seg.get(u"id", -1)) if seg is not None else -1
        except Exception:
            seg_id_i = -1
        seg_sel = (
            seg_id_i in selected_set
            and focus == u"tramo"
        )
        try:
            seg_primary = seg_sel and seg_id_i == int(selected_segment)
        except Exception:
            seg_primary = seg_sel
        wall_sel = wi in sel_walls and focus == u"muro"
        wall_primary = wall_sel and wi == sel_wall
        is_base = wi == base_i
        emp = False
        if not is_base:
            emp = effective_empalme(
                modes[wi] if wi < len(modes) else TROCEO_AUTO,
                autos[wi] if wi < len(autos) else False,
            )
        try:
            thick_key = int(round(float(m.get(u"thick_mm") or 200)))
        except Exception:
            thick_key = 200
        hx = color_mp.get(thick_key, _THICKNESS_UI_PALETTE[0])
        x_off, draw_w = _draw_w_xoff(
            m, elev_col_w, scale_px, g_min=lay_g_min, span_u=lay_span_u,
        )
        fund_h_mm = _wall_fund_height_mm(m)
        fund_h_px = max(0.0, fund_h_mm * scale_px) if fund_h_mm > 0.1 else 0.0
        wall_h_px = max(8.0, row_h - fund_h_px)
        stem_foot_y = float(y + wall_h_px)
        fund_info = m.get(u"fund_info")
        is_first = di <= 0
        is_last = di >= n - 1
        conf = confs[wi] if wi < len(confs) else {}
        mesh = meshes[wi] if wi < len(meshes) else {}

        bg = Border()
        bg.Width = elev_col_w
        bg.Height = row_h
        bg.Background = br_app
        bg.BorderThickness = Thickness(0)
        Canvas.SetLeft(bg, elev_x)
        Canvas.SetTop(bg, y)
        canvas.Children.Add(bg)

        if is_first and is_last:
            band_bt = Thickness(1.0)
        elif is_first:
            band_bt = Thickness(1.0, 1.0, 1.0, 0.0)
        elif is_last:
            band_bt = Thickness(1.0, 0.0, 1.0, 1.0)
        else:
            band_bt = Thickness(1.0, 0.0, 1.0, 0.0)

        wall_band = Border()
        wall_band.Width = draw_w
        wall_band.Height = wall_h_px
        if wall_sel:
            wall_band.Background = _hex_brush("#4ade80", 56 if wall_primary else 34)
            wall_band.BorderBrush = br_wall
            wall_band.BorderThickness = Thickness(2.0 if wall_primary else 1.5)
        elif seg_sel:
            wall_band.Background = _hex_brush(
                "#5BC0DE", 48 if seg_primary else 28
            )
            wall_band.BorderBrush = br_sel
            wall_band.BorderThickness = Thickness(1.5 if seg_primary else 1.25)
        else:
            wall_band.Background = _hex_brush(hx, 34)
            wall_band.BorderBrush = _hex_brush(hx, 170)
            wall_band.BorderThickness = band_bt
        try:
            wall_band.SnapsToDevicePixels = True
        except Exception:
            pass
        Canvas.SetLeft(wall_band, elev_x + x_off)
        Canvas.SetTop(wall_band, y)

        sp = StackPanel()
        sp.HorizontalAlignment = HorizontalAlignment.Center
        sp.VerticalAlignment = VerticalAlignment.Center
        tipo = _fill_wall_fuste_label(
            sp,
            m,
            thick_key,
            conf,
            mesh,
            show_mesh,
            br_title,
            br_wall,
            br_muted,
            br_mesh_accent,
            hx,
            TextBlock,
            StackPanel,
            FontWeights,
            HorizontalAlignment,
            VerticalAlignment,
            TextAlignment,
            Thickness,
            Orientation,
        )
        wall_band.Child = sp

        tip = ToolTip()
        try:
            len_mm = int(round(float(m.get(u"length_ft") or m.get(u"length_u") or 0.0) * 304.8))
        except Exception:
            len_mm = 0
        if show_mesh:
            mesh_on = True
            try:
                if u"activo" in (mesh or {}):
                    mesh_on = bool(mesh.get(u"activo"))
            except Exception:
                mesh_on = True
            if mesh_on:
                tip.Content = (
                    u"{0} · {1} · e={2} · L={3}\n"
                    u"D.M.  V.={4}  H.={5}\n"
                    u"Clic = seleccionar · Ctrl+clic = multi · Mayús+clic = rango"
                ).format(
                    m.get(u"name") or u"W{0}".format(wi),
                    tipo,
                    thick_key,
                    len_mm,
                    _mesh_bar_spec_txt(
                        mesh.get(u"v_diam_mm"), mesh.get(u"v_spacing_mm"),
                    ),
                    _mesh_bar_spec_txt(
                        mesh.get(u"h_diam_mm"), mesh.get(u"h_spacing_mm"),
                    ),
                )
            else:
                tip.Content = (
                    u"{0} · {1} · e={2} · L={3}\n"
                    u"sin malla\n"
                    u"Clic = seleccionar · Ctrl+clic = multi · Mayús+clic = rango"
                ).format(
                    m.get(u"name") or u"W{0}".format(wi),
                    tipo,
                    thick_key,
                    len_mm,
                )
        else:
            tip.Content = (
                u"{0} · {1} · e={2} · L={3}\n"
                u"Clic = seleccionar · Ctrl+clic = multi · Mayús+clic = rango"
            ).format(
                m.get(u"name") or u"W{0}".format(wi),
                tipo,
                thick_key,
                len_mm,
            )
        wall_band.ToolTip = tip

        try:

            def _mk_wall(idx):
                def _h(sender, args):
                    if on_select_wall is not None:
                        on_select_wall(int(idx))

                return _h

            wall_band.MouseLeftButtonDown += MouseButtonEventHandler(_mk_wall(wi))
        except Exception:
            pass

        canvas.Children.Add(wall_band)
        y_bottom_of_wall[wi] = stem_foot_y
        band_geom[wi] = (elev_x + x_off, draw_w, y)

        if fund_h_px > 0.5:
            try:
                _draw_foundation_footing_elev(
                    canvas,
                    elev_x,
                    x_off,
                    draw_w,
                    stem_foot_y,
                    scale_px,
                    fund_h_mm,
                    _wall_fund_width_mm(m),
                    hx,
                    fund_info,
                    Canvas,
                    TextBlock,
                    FontWeights,
                )
            except Exception:
                pass

        if on_draw_extremo_marks is not None:
            try:
                on_draw_extremo_marks(
                    canvas, int(wi), float(elev_x + x_off), float(draw_w),
                    float(y), float(wall_h_px),
                )
            except Exception:
                pass

        # Diferir marca de nivel (texto+disco+guía): pie del fuste, no bajo fundación.
        level_marks.append(
            (stem_foot_y, _level_label_text(m))
        )

        def _draw_pie_ctrl(ctrl_x, modes_side, on_cycle, side_active, extremo_lbl):
            try:
                opac = float(dim_op if not side_active else 1.0)
            except Exception:
                opac = 1.0 if side_active else BAND_DIM_OPACITY
            opac = max(0.15, min(1.0, opac))
            if is_base:
                tb = TextBlock()
                tb.Text = u"Base"
                tb.Foreground = br_muted
                tb.FontSize = 10.0
                try:
                    tb.Opacity = opac
                except Exception:
                    pass
                Canvas.SetLeft(tb, ctrl_x)
                Canvas.SetTop(tb, y + row_h - 18.0)
                canvas.Children.Add(tb)
                return
            mode = modes_side[wi] if wi < len(modes_side) else TROCEO_AUTO
            auto = autos[wi] if wi < len(autos) else False
            emp_side = effective_empalme(mode, auto)
            cap = pie_caption(mode, auto)
            manual = mode != TROCEO_AUTO
            btn = Button()
            btn.Content = cap
            btn.Width = BTN_W
            btn.Height = BTN_H
            btn.FontSize = 11.0
            btn.FontWeight = FontWeights.SemiBold
            btn.Padding = Thickness(4.0, 1.0, 4.0, 1.0)
            btn.BorderThickness = Thickness(1.0)
            try:
                btn.Opacity = opac
            except Exception:
                pass
            if manual:
                btn.Background = br_manual_bg
                btn.BorderBrush = br_accent
                btn.Foreground = br_accent
            elif emp_side:
                btn.Background = _hex_brush("#3f1515")
                btn.BorderBrush = br_red
                btn.Foreground = br_red
            else:
                btn.Background = br_app
                btn.BorderBrush = br_border
                btn.Foreground = br_muted
            tip2 = ToolTip()
            tip2.Content = (
                u"{0}: ciclo Auto → Tramo → Cont. (independiente por extremo)"
            ).format(extremo_lbl)
            btn.ToolTip = tip2
            Canvas.SetLeft(btn, ctrl_x)
            Canvas.SetTop(btn, y + row_h - BTN_H - 4.0)

            def _mk_cycle(idx, cb):
                def _h(sender, args):
                    if cb is not None:
                        cb(int(idx))

                return _h

            try:
                btn.Click += RoutedEventHandler(_mk_cycle(wi, on_cycle))
            except Exception:
                btn.Click += _mk_cycle(wi, on_cycle)
            canvas.Children.Add(btn)

        if dual_bands:
            _draw_pie_ctrl(
                ctrl_x_left, side_left[u"modes"], side_left[u"cycle"],
                act == ex_left, _extremo_side_label(ex_left),
            )
            _draw_pie_ctrl(
                ctrl_x_right, side_right[u"modes"], side_right[u"cycle"],
                act == ex_right, _extremo_side_label(ex_right),
            )
        else:
            _draw_pie_ctrl(
                ctrl_x_right, modes, on_cycle_wall,
                True, u"Tramo",
            )

        if emp:
            tick = Rectangle()
            tick.Width = 12.0
            tick.Height = 3.0
            tick.Fill = br_red
            Canvas.SetLeft(tick, elev_x + x_off + 2.0)
            Canvas.SetTop(tick, stem_foot_y - 2.0)
            canvas.Children.Add(tick)

    lap_mm = lap_mm_from_long_diam(long_diam_mm)
    lap_px = max(8.0, float(lap_mm) * scale_px)
    for wi in empalme_indices_from_modes(modes, autos, base_index=base_i):
        geom = band_geom.get(wi)
        y_foot = y_bottom_of_wall.get(wi)
        if geom is None or y_foot is None:
            continue
        bx, bw, _yt = geom
        ln_a = Line()
        ln_a.X1 = bx
        ln_a.X2 = bx + bw
        ln_a.Y1 = y_foot
        ln_a.Y2 = y_foot
        ln_a.Stroke = br_a
        ln_a.StrokeThickness = 2.0
        canvas.Children.Add(ln_a)
        ln_b = Line()
        ln_b.X1 = bx
        ln_b.X2 = bx + bw
        ln_b.Y1 = y_foot - lap_px
        ln_b.Y2 = y_foot - lap_px
        ln_b.Stroke = br_b
        ln_b.StrokeThickness = 2.0
        try:
            from System.Windows.Media import DoubleCollection

            dc = DoubleCollection()
            dc.Add(4.0)
            dc.Add(3.0)
            ln_b.StrokeDashArray = dc
        except Exception:
            pass
        canvas.Children.Add(ln_b)

    # Cotas de nivel encima de bg/empalme para que no las tape la fila inferior.
    for foot_y, lbl in level_marks:
        try:
            _draw_level_marker(canvas, elev_x, elev_col_w, foot_y, lbl)
        except Exception:
            pass
