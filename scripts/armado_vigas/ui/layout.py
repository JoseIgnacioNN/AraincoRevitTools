# -*- coding: utf-8 -*-
"""Layout horizontal del canvas (port del mockup HTML)."""

from __future__ import division

BEAM_GAP_PX = 8.0
CANVAS_SIDE_PAD_PX = 12.0
BEAM_SLOT_PX = 360.0  # fallback cuando no hay longitud conocida
MIN_BEAM_SLOT_PX = 168.0
# Solo referencia legacy; el alzado ya no “encaja” al viewport.
FIT_MIN_SLOT_PX = 88.0
TRAMO_PANEL_MAX_PX = 360.0
TRAMO_PANEL_FLOOR_PX = 228.0
# Legacy: panel dual sup+inf (ya no usado en canvas).
TRAMO_PANEL_CONTENT_PX = 262.0
ESTRIBO_PANEL_MIN_PX = 120.0
ESTRIBO_PANEL_MAX_PX = 280.0

# Escala base del alzado (px por unidad de modelo). U y V isótropas.
# Si el dibujo cabe en el viewport, se aplica zoom-extents (solo ampliar).
MODEL_PX_PER_MM = 0.08
MODEL_PX_PER_FT = MODEL_PX_PER_MM * 304.8  # ≈ 24.384
# Límite superior del zoom-extents (evita siluetas ridículas en lotes muy cortos).
MODEL_PX_PER_FT_MAX = MODEL_PX_PER_FT * 6.0
ZOOM_EXTENTS_MARGIN = 0.96  # relleno útil del viewport (aire en bordes)

# Zoom interactivo del canvas de elevación (sobre la escala del layout).
VIEW_ZOOM_DEFAULT = 1.0
VIEW_ZOOM_MIN = 0.25
VIEW_ZOOM_MAX = 4.0
VIEW_ZOOM_STEP = 1.15
_FT_PER_M = 3.280839895013123

ELEVATION_HEIGHT_PX = 200.0  # fallback sin datos de elevación del modelo
ELEVATION_HEIGHT_MIN_PX = 80.0
ELEVATION_PAD_PX = 10.0
LABELS_HEIGHT_PX = 44.0
ZONE_PANEL_LABEL_PX = 38.0
ZONE_PANEL_ROW_PX = 26.0
ZONE_PANEL_FOOTER_PX = 16.0
SUPLE_LABEL_PX = 52.0
ESTRIBO_SECTION_HDR_PX = 16.0
ESTRIBO_PAIR_BODY_PX = 88.0
CONFIN_INLINE_ROW_PX = 30.0
SUPLE_INF_ROW_PX = 86.0
SUPLE_SUP_ROW_PX = 108.0
SUPLE_SLOT_HDR_PX = 18.0
ESTRIBO_SLOT_HDR_PX = 18.0
ESTRIBO_SLOT_VPAD_PX = 18.0
SECTION_RAIL_PAD_PX = 16.0
SECTION_RAIL_WIDTH_PX = 340.0
SECTION_CTRL_WIDTH_PX = SECTION_RAIL_WIDTH_PX - SECTION_RAIL_PAD_PX
SUPLE_CANVAS_SLOT_HEIGHT_PX = (
    SUPLE_SLOT_HDR_PX
    + SUPLE_INF_ROW_PX
    + 10.0
)
SUPLE_SUP_CANVAS_SLOT_HEIGHT_PX = (
    SUPLE_SLOT_HDR_PX
    + SUPLE_SUP_ROW_PX
    + 10.0
)
ESTRIBO_ZONE_HEIGHT_PX = SUPLE_CANVAS_SLOT_HEIGHT_PX
# Bandas Tn — pill multi-línea (opción C): T n + n×ø por capa, trazo fino
# Alturas calibradas a métricas reales WPF (FontSize 10 ≈ ~15 px de línea + padding).
TRAMO_BAND_CTRL_HEIGHT_PX = 48.0
TRAMO_BAND_TINT_HEIGHT_PX = 2.0  # legacy alias; A usa TRAZO
TRAMO_BAND_TRAZO_H_PX = 2.0
TRAMO_BAND_TRAZO_H_SEL_PX = 3.0
TRAMO_BAND_PILL_TITLE_FONT_PX = 10.0
TRAMO_BAND_PILL_LAYER_FONT_PX = 10.0
TRAMO_BAND_PILL_FONT_PX = 10.0  # alias título
TRAMO_BAND_PILL_TITLE_H_PX = 16.0
TRAMO_BAND_PILL_LAYER_H_PX = 15.0
TRAMO_BAND_PILL_TITLE_GAP_PX = 2.0
# padding vertical Border (4+5) + 1 px borde x2
TRAMO_BAND_PILL_PAD_V_PX = 11.0
TRAMO_BAND_TRAZO_SLOT_PX = 7.0
TRAMO_BAND_PILL_SLACK_PX = 2.0  # holgura de layout (anti-recorte)
# Legacy: alto mín. 1 capa (título + 1 fila n×ø)
TRAMO_BAND_PILL_H_PX = (
    TRAMO_BAND_PILL_PAD_V_PX
    + TRAMO_BAND_PILL_TITLE_H_PX
    + TRAMO_BAND_PILL_TITLE_GAP_PX
    + TRAMO_BAND_PILL_LAYER_H_PX
    + TRAMO_BAND_PILL_SLACK_PX
)
TRAMO_BAND_COLLAPSED_HEIGHT_PX = TRAMO_BAND_PILL_H_PX + TRAMO_BAND_TRAZO_SLOT_PX
TRAMO_BAND_LAYER_ROW_PX = 22.0
TRAMO_BAND_CAP_ROW_PX = 24.0
TRAMO_BAND_VPAD_PX = 1.0
TRAMO_BAND_HEIGHT_PX = 16.0
TRAMO_BAND_MAX_CAPAS = 3
# Empalme por viga (pill encima/bajo silueta) — define troceo de Tn
TRAMO_EMPALME_PILL_H_PX = 22.0
TRAMO_EMPALME_PILL_FONT_PX = 10.0
TRAMO_EMPALME_ROW_PX = 26.0
TRAMO_EMPALME_BTN_ROW_PX = TRAMO_EMPALME_ROW_PX
TRAMO_PANEL_LANE_PX = 148.0
TRAMO_PANEL_W_PX = 172.0
LANE_GAP_PX = 8.0
FACE_BLOCK_PAD_PX = 12.0
TRAMO_CTRL_HEIGHT_PX = 172.0
TRAMO_CTRL_HEIGHT_SINGLE_PX = 118.0
TRAMO_FACE_HDR_PX = 18.0
TRAMO_FACE_ZONE_PAD_PX = 4.0
# Ancho panel Tn por cara (alias del mockup --panel-w)
TRAMO_PANEL_SINGLE_FACE_PX = TRAMO_PANEL_W_PX
AXIS_HINT_HEIGHT_PX = 22.0


def tramo_band_pill_height_px(n_capas):
    """Alto del chip multi-línea (título + una fila por capa), con holgura WPF."""
    try:
        n = int(n_capas or 1)
    except (TypeError, ValueError):
        n = 1
    n = max(1, min(int(TRAMO_BAND_MAX_CAPAS), n))
    return float(
        TRAMO_BAND_PILL_PAD_V_PX
        + TRAMO_BAND_PILL_TITLE_H_PX
        + TRAMO_BAND_PILL_TITLE_GAP_PX
        + float(n) * TRAMO_BAND_PILL_LAYER_H_PX
        + TRAMO_BAND_PILL_SLACK_PX
    )


def tramo_band_body_height_px(selected, n_capas):
    """Altura cuerpo banda Tn (pill multi-línea + trazo)."""
    return float(tramo_band_cell_height_px(selected, n_capas))


def tramo_band_cell_height_px(selected, n_capas):
    """Alto total selector Tn (pill multi-línea + hueco de trazo)."""
    return float(tramo_band_pill_height_px(n_capas) + TRAMO_BAND_TRAZO_SLOT_PX)


def _beam_len_m(beam):
    try:
        return max(0.0, float((beam or {}).get("len") or 0.0))
    except (TypeError, ValueError):
        return 0.0


def compute_layout(sorted_beams, viewport_w, apoyos=None, use_model_positions=False, joined=None, viewport_h=None):
    # Siempre preferir posiciones del modelo si hay uStart/uEnd (vista activa).
    joined_list = _joined_list_from_session_or_arg(joined)
    if use_model_positions:
        has_model_u = any(
            b.get("uStart") is not None and b.get("uEnd") is not None
            for b in (sorted_beams or [])
        )
        if has_model_u:
            model_result = compute_layout_model(
                sorted_beams, viewport_w, apoyos or [], joined=joined_list,
                viewport_h=viewport_h,
            )
            if model_result.get("modelPositions"):
                return model_result
        elif apoyos:
            model_result = compute_layout_model(
                sorted_beams, viewport_w, apoyos, joined=joined_list,
                viewport_h=viewport_h,
            )
            if model_result.get("modelPositions"):
                return model_result

    n = len(sorted_beams or [])
    viewport_w = max(320.0, float(viewport_w or 640.0))
    if not n:
        return {"layouts": [], "contentWidthPx": viewport_w, "needsScroll": False}

    gaps_total = BEAM_GAP_PX * max(n - 1, 0)
    lengths = [_beam_len_m(b) for b in sorted_beams]
    # Anchos proporcional a longitud real (escala fija); no se rellenan al viewport.
    px_per_m = float(MODEL_PX_PER_FT) * float(_FT_PER_M)
    widths = []
    for length_m in lengths:
        if length_m > 1e-9:
            widths.append(max(4.0, float(length_m) * px_per_m))
        else:
            widths.append(float(BEAM_SLOT_PX))
    content_w = CANVAS_SIDE_PAD_PX * 2.0 + gaps_total + sum(widths)
    needs_scroll = content_w > viewport_w + 1.0

    cursor = CANVAS_SIDE_PAD_PX
    layouts = []
    for idx, beam in enumerate(sorted_beams):
        width_px = float(widths[idx])
        left_px = cursor
        center_px = left_px + width_px * 0.5
        cursor += width_px + BEAM_GAP_PX
        layouts.append({
            "idx": idx,
            "leftPx": left_px,
            "widthPx": width_px,
            "centerPx": center_px,
            "leftPct": (left_px / content_w) * 100.0 if content_w else 0.0,
            "widthPct": (width_px / content_w) * 100.0 if content_w else 0.0,
            "centerPct": (center_px / content_w) * 100.0 if content_w else 0.0,
        })
    return {
        "layouts": layouts,
        "contentWidthPx": content_w,
        "needsScroll": needs_scroll,
        "modelPositions": False,
        "pxPerFtU": float(MODEL_PX_PER_FT),
        "pxPerFtV": float(MODEL_PX_PER_FT),
        "elevPadPx": float(ELEVATION_PAD_PX),
        "elevHeightPx": float(ELEVATION_HEIGHT_PX),
    }


def _joined_list_from_session_or_arg(joined):
    """Acepta lista de records o dict tipo SESSION.joined_framing."""
    if not joined:
        return []
    if isinstance(joined, dict):
        return list(joined.get("all") or [])
    return list(joined)


def model_u_span(sorted_beams, apoyos, joined=None):
    """Rango ``[u_min, u_max]`` proyectado sobre ``view.RightDirection``."""
    values = []
    for beam in sorted_beams or []:
        for key in ("uStart", "uEnd", "solidUMin", "solidUMax"):
            v = beam.get(key)
            if v is None:
                continue
            try:
                values.append(float(v))
            except (TypeError, ValueError):
                pass
    for ap in apoyos or []:
        for key in ("uMin", "uMax", "uView"):
            v = ap.get(key)
            if v is None:
                continue
            try:
                values.append(float(v))
            except (TypeError, ValueError):
                pass
        # Mitad del ancho proyectado alrededor del centro si solo hay uView + width
        try:
            u = ap.get("uView")
            w_mm = ap.get("widthMm") or ap.get("thicknessMm")
            if u is not None and w_mm:
                half = (float(w_mm) / 304.8) * 0.5
                values.append(float(u) - half)
                values.append(float(u) + half)
        except (TypeError, ValueError):
            pass
    for rec in joined or []:
        for key in ("uStart", "uEnd", "solidUMin", "solidUMax", "uMin", "uMax"):
            v = rec.get(key)
            if v is None:
                continue
            try:
                values.append(float(v))
            except (TypeError, ValueError):
                pass
    if not values:
        return None
    u_min = min(values)
    u_max = max(values)
    if u_max - u_min < 1e-9:
        u_max = u_min + 1e-9
    return u_min, u_max


def model_v_span(sorted_beams, apoyos, joined=None):
    """Rango ``[v_min, v_max]`` proyectado sobre ``view.UpDirection``."""
    values = []
    for beam in sorted_beams or []:
        for key in ("vMin", "vMax"):
            v = beam.get(key)
            if v is None:
                continue
            try:
                values.append(float(v))
            except (TypeError, ValueError):
                pass
    for ap in apoyos or []:
        for key in ("vMin", "vMax"):
            v = ap.get(key)
            if v is None:
                continue
            try:
                values.append(float(v))
            except (TypeError, ValueError):
                pass
    for rec in joined or []:
        for key in ("vMin", "vMax"):
            v = rec.get(key)
            if v is None:
                continue
            try:
                values.append(float(v))
            except (TypeError, ValueError):
                pass
    if not values:
        return None
    v_min = min(values)
    v_max = max(values)
    if v_max - v_min < 1e-9:
        v_max = v_min + 1e-9
    return v_min, v_max


def _model_scale(content_w, u_min, u_max):
    u_range = max(float(u_max) - float(u_min), 1e-9)
    pool = float(content_w) - CANVAS_SIDE_PAD_PX * 2.0
    return max(pool, 1.0) / u_range


def model_u_to_left_pct(u, u_min, u_max, content_w):
    """Convierte escalar de vista a ``leftPct`` coherente con :func:`compute_layout_model`."""
    scale = _model_scale(content_w, u_min, u_max)
    left_px = CANVAS_SIDE_PAD_PX + (float(u) - float(u_min)) * scale
    cw = float(content_w or 1.0)
    return (left_px / cw) * 100.0


def _layout_from_model_u(sorted_beams, u_min, u_max, scale, content_w):
    layouts = []
    cw = float(content_w or 1.0)
    for idx, beam in enumerate(sorted_beams):
        try:
            u0 = float(beam.get("uStart", u_min))
            u1 = float(beam.get("uEnd", u0))
        except (TypeError, ValueError):
            u0, u1 = float(u_min), float(u_max)
        if u1 < u0:
            u0, u1 = u1, u0
        left_px = CANVAS_SIDE_PAD_PX + (u0 - float(u_min)) * scale
        width_px = max(4.0, (u1 - u0) * scale)
        center_px = left_px + width_px * 0.5
        layouts.append({
            "idx": idx,
            "leftPx": left_px,
            "widthPx": width_px,
            "centerPx": center_px,
            "leftPct": (left_px / cw) * 100.0,
            "widthPct": (width_px / cw) * 100.0,
            "centerPct": (center_px / cw) * 100.0,
        })
    return layouts


def _elev_stage_chrome_height_px():
    """Alto del stage de alzado fuera del bloque V (bandas max-capa + cabecera + labels/eje)."""
    # Reserva peor caso multi-capa para no recortar pills al hacer zoom extents.
    band = float(tramo_band_cell_height_px(False, TRAMO_BAND_MAX_CAPAS))
    return (
        band * 2.0
        + 22.0  # cabecera alzado
        + 6.0   # márgenes stage
        + float(LABELS_HEIGHT_PX)
        + float(AXIS_HINT_HEIGHT_PX)
        + 8.0   # margen inferior
    )


def compute_layout_model(sorted_beams, viewport_w, apoyos, joined=None, viewport_h=None):
    """Layout modelo isótropo (U=V) con zoom-extents si cabe en el viewport.

    - Escala base ``MODEL_PX_PER_FT`` (proporciones reales).
    - Si el dibujo a escala base es menor que el espacio útil, se amplía la misma
      escala U=V hasta llenar el viewport (solo ampliar, no comprimir).
    - Si no cabe, se conserva la escala base y hay scroll.
    """
    n = len(sorted_beams or [])
    viewport_w = max(320.0, float(viewport_w or 640.0))
    joined_list = _joined_list_from_session_or_arg(joined)
    span = model_u_span(sorted_beams, apoyos, joined=joined_list)
    if not n or span is None:
        return {"modelPositions": False}

    u_min, u_max = span
    u_range = max(float(u_max) - float(u_min), 1e-9)
    elev_pad = float(ELEVATION_PAD_PX)

    v_span_pair = model_v_span(sorted_beams, apoyos, joined=joined_list)
    if v_span_pair is not None:
        v_min, v_max = v_span_pair
        v_span = max(float(v_max) - float(v_min), 1e-9)
    else:
        v_min = 0.0
        base = float(MODEL_PX_PER_FT)
        v_span = max(float(ELEVATION_HEIGHT_PX) - elev_pad * 2.0, 24.0) / base
        v_max = v_min + v_span

    base_scale = float(MODEL_PX_PER_FT)
    scale = base_scale
    zoom_extents = False

    try:
        vp_h = float(viewport_h) if viewport_h is not None else 0.0
    except (TypeError, ValueError):
        vp_h = 0.0

    if vp_h > 80.0 and viewport_w > 80.0:
        pad_u = float(CANVAS_SIDE_PAD_PX) * 2.0
        avail_u = max(32.0, float(viewport_w) - pad_u)
        scale_u = avail_u / u_range

        chrome = _elev_stage_chrome_height_px()
        avail_elev = max(float(ELEVATION_HEIGHT_MIN_PX), vp_h - chrome)
        avail_v_body = max(16.0, avail_elev - elev_pad * 2.0)
        scale_v = avail_v_body / v_span

        fit = min(scale_u, scale_v) * float(ZOOM_EXTENTS_MARGIN)
        # Solo ampliar: si el lote es "pequeño" en pantalla, acercar (zoom extents).
        if fit > base_scale:
            scale = min(float(MODEL_PX_PER_FT_MAX), fit)
            zoom_extents = scale > base_scale * 1.02

    content_w = float(CANVAS_SIDE_PAD_PX) * 2.0 + u_range * scale
    needs_scroll = content_w > viewport_w + 1.0

    layouts = _layout_from_model_u(sorted_beams, u_min, u_max, scale, content_w)

    elev_h = max(
        float(ELEVATION_HEIGHT_MIN_PX),
        elev_pad * 2.0 + v_span * scale,
    )
    if vp_h > 80.0 and not needs_scroll:
        # Con zoom extents, el bloque de alzado puede absorber altura sobrante mínima
        # del viewport sin anisotropía (misma escala, solo marco).
        chrome = _elev_stage_chrome_height_px()
        target = max(elev_h, min(vp_h - chrome, elev_h * 1.15))
        if zoom_extents:
            elev_h = max(elev_h, min(target, vp_h - chrome))

    return {
        "layouts": layouts,
        "contentWidthPx": content_w,
        "needsScroll": needs_scroll,
        "modelPositions": True,
        "modelUMin": u_min,
        "modelUMax": u_max,
        "modelVMin": v_min,
        "modelVMax": v_max,
        "pxPerFtU": scale,
        "pxPerFtV": scale,
        "elevPadPx": elev_pad,
        "elevHeightPx": elev_h,
        "zoomExtents": zoom_extents,
        "basePxPerFt": base_scale,
    }


def tramo_span(layouts, tramo, content_width_px):
    """Span visual del Tn en % del canvas (izquierda → derecha).

    ``edgeStart`` / ``edgeEnd`` aplican a vigas extremas del tramo en orden de
    cadena; el tramo se ancla por ``leftPct`` mínimo (nunca asume que el primer
    índice de lista es el más a la izquierda si el layout lo contradice).
    """
    idxs = list(tramo.get("beamIndices") or [])
    if not idxs or not layouts:
        return {
            "leftPct": 0.0,
            "widthPct": 0.0,
            "centerPct": 0.0,
            "widthPx": 0.0,
        }
    # Extremos de cadena (troceo empalme) vs extremos visuales.
    i_chain0 = int(idxs[0])
    i_chain1 = int(idxs[-1])
    valid = [i for i in idxs if 0 <= int(i) < len(layouts)]
    if not valid:
        return {
            "leftPct": 0.0,
            "widthPct": 0.0,
            "centerPct": 0.0,
            "widthPx": 0.0,
        }
    by_left = sorted(valid, key=lambda i: float(layouts[int(i)].get("leftPct") or 0.0))
    i_vis_left = int(by_left[0])
    i_vis_right = int(by_left[-1])
    left_lay = layouts[i_vis_left]
    right_lay = layouts[i_vis_right]
    left_pct = float(left_lay["leftPct"])
    right_pct = float(right_lay["leftPct"]) + float(right_lay["widthPct"])

    # half @ start de cadena: recorta el side del tramo que corresponde a esa viga.
    if tramo.get("edgeStart") == "half" and 0 <= i_chain0 < len(layouts):
        lay0 = layouts[i_chain0]
        mid0 = float(lay0["leftPct"]) + float(lay0["widthPct"]) * 0.5
        # Si la viga de edgeStart es el extremo visual izquierdo, sube left; si no, baja right.
        if i_chain0 == i_vis_left:
            left_pct = mid0
        elif i_chain0 == i_vis_right:
            right_pct = mid0
        else:
            # edgeStart en media: recortar hacia el interior del tramo.
            left_pct = max(left_pct, mid0) if mid0 < (left_pct + right_pct) * 0.5 else left_pct
            right_pct = min(right_pct, mid0) if mid0 >= (left_pct + right_pct) * 0.5 else right_pct
    if tramo.get("edgeEnd") == "half" and 0 <= i_chain1 < len(layouts):
        lay1 = layouts[i_chain1]
        mid1 = float(lay1["leftPct"]) + float(lay1["widthPct"]) * 0.5
        if i_chain1 == i_vis_right:
            right_pct = mid1
        elif i_chain1 == i_vis_left:
            left_pct = mid1
        else:
            left_pct = max(left_pct, mid1) if mid1 < (left_pct + right_pct) * 0.5 else left_pct
            right_pct = min(right_pct, mid1) if mid1 >= (left_pct + right_pct) * 0.5 else right_pct

    if right_pct < left_pct:
        left_pct, right_pct = right_pct, left_pct
    width_pct = right_pct - left_pct
    content_w = float(content_width_px or 1.0)
    return {
        "leftPct": left_pct,
        "widthPct": width_pct,
        "centerPct": (left_pct + right_pct) * 0.5,
        "widthPx": (width_pct / 100.0) * content_w,
    }


def panel_width_for_slot(slot_px, max_px, floor_px=None):
    floor_px = floor_px if floor_px is not None else 72.0
    return min(float(max_px), max(float(floor_px), float(slot_px) - 6.0))


def beam_canvas_label(idx):
    """Numeración UI en canvas: izquierda → derecha = Viga 1…N."""
    return u"Viga {0}".format(int(idx) + 1)


def pct_to_px(pct, content_w):
    return (float(pct) / 100.0) * float(content_w)


def build_support_chain(sorted_beams, layouts, apoyos=None, layout_meta=None):
    if layout_meta and layout_meta.get("modelPositions") and apoyos:
        chain = _build_support_chain_model(sorted_beams, apoyos, layout_meta)
        if chain:
            return chain
    return _build_support_chain_from_beams(sorted_beams, layouts)


def _build_support_chain_from_beams(sorted_beams, layouts):
    chain = []
    last_id = None
    for i, beam in enumerate(sorted_beams):
        lay = layouts[i]
        points = [
            {"id": beam.get("colStart"), "pct": lay["leftPct"]},
            {"id": beam.get("colEnd"), "pct": lay["leftPct"] + lay["widthPct"]},
        ]
        for p in points:
            pid = p.get("id")
            if pid and pid != last_id:
                chain.append(p)
                last_id = pid
    return chain


def _build_support_chain_model(sorted_beams, apoyos, layout_meta):
    u_min = layout_meta.get("modelUMin")
    u_max = layout_meta.get("modelUMax")
    content_w = layout_meta.get("contentWidthPx", 1.0)
    if u_min is None or u_max is None:
        return []

    # Cadena = apoyos verticales ordenados por U (columnas/muros; sin losas).
    items = []
    seen = set()
    for ap in apoyos or []:
        try:
            kind = unicode(ap.get("kind") or u"").lower()
        except Exception:
            kind = u""
        if kind in (u"floor", u"losa", u"slab"):
            continue
        aid = ap.get("id")
        if not aid or aid in seen:
            continue
        u = ap.get("uView")
        if u is None:
            umin, umax = ap.get("uMin"), ap.get("uMax")
            if umin is not None and umax is not None:
                u = (float(umin) + float(umax)) * 0.5
        if u is None:
            continue
        seen.add(aid)
        items.append({"id": aid, "u": float(u)})
    if not items:
        # Fallback: extremos referenciados por vigas
        apoyo_by_id = {a["id"]: a for a in (apoyos or []) if a.get("id")}
        last_id = None
        for beam in sorted_beams:
            for key in ("colStart", "colEnd"):
                aid = beam.get(key)
                if not aid or aid == last_id:
                    continue
                ap = apoyo_by_id.get(aid)
                try:
                    kind = unicode((ap or {}).get("kind") or u"").lower()
                except Exception:
                    kind = u""
                if kind in (u"floor", u"losa", u"slab"):
                    continue
                u = ap.get("uView") if ap else None
                if u is None:
                    continue
                items.append({"id": aid, "u": float(u)})
                last_id = aid
    items.sort(key=lambda it: it["u"])
    chain = []
    for it in items:
        pct = model_u_to_left_pct(it["u"], u_min, u_max, content_w)
        chain.append({"id": it["id"], "pct": pct})
    return chain


def collect_apoyos(sorted_beams):
    ids = set()
    for beam in sorted_beams or []:
        if beam.get("colStart"):
            ids.add(beam["colStart"])
        if beam.get("colEnd"):
            ids.add(beam["colEnd"])
    cols = 0
    walls = 0
    for aid in ids:
        if unicode(aid).startswith(u"M"):
            walls += 1
        else:
            cols += 1
    return {"ids": sorted(ids), "cols": cols, "walls": walls}
