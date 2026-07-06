# -*- coding: utf-8 -*-
"""Layout horizontal del canvas (port del mockup HTML)."""

from __future__ import division

BEAM_GAP_PX = 8.0
CANVAS_SIDE_PAD_PX = 12.0
BEAM_SLOT_PX = 360.0  # fallback cuando no hay longitud conocida
MIN_BEAM_SLOT_PX = 168.0
FIT_MIN_SLOT_PX = 88.0
TRAMO_PANEL_MAX_PX = 360.0
TRAMO_PANEL_FLOOR_PX = 228.0
# Legacy: panel dual sup+inf (ya no usado en canvas).
TRAMO_PANEL_CONTENT_PX = 262.0
ESTRIBO_PANEL_MIN_PX = 120.0
ESTRIBO_PANEL_MAX_PX = 280.0

ELEVATION_HEIGHT_PX = 200.0
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
# Opción D — bandas Tn con controles alineadas al alzado
TRAMO_BAND_CTRL_HEIGHT_PX = 48.0
TRAMO_BAND_TINT_HEIGHT_PX = 6.0
TRAMO_BAND_COLLAPSED_HEIGHT_PX = 28.0
TRAMO_BAND_LAYER_ROW_PX = 22.0
TRAMO_BAND_CAP_ROW_PX = 24.0
TRAMO_BAND_VPAD_PX = 4.0
TRAMO_BAND_HEIGHT_PX = 16.0
TRAMO_EMPALME_ROW_PX = 36.0
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


def tramo_band_body_height_px(selected, n_capas):
    """Altura cuerpo banda Tn (sin franja tinte)."""
    if not selected:
        return TRAMO_BAND_COLLAPSED_HEIGHT_PX
    n = max(1, int(n_capas or 1))
    return TRAMO_BAND_VPAD_PX * 2 + TRAMO_BAND_CAP_ROW_PX + n * TRAMO_BAND_LAYER_ROW_PX


def tramo_band_cell_height_px(selected, n_capas):
    return tramo_band_body_height_px(selected, n_capas) + TRAMO_BAND_TINT_HEIGHT_PX


def _beam_len_m(beam):
    try:
        return max(0.0, float((beam or {}).get("len") or 0.0))
    except (TypeError, ValueError):
        return 0.0


def compute_layout(sorted_beams, viewport_w, apoyos=None, use_model_positions=False):
    if use_model_positions and apoyos:
        model_result = compute_layout_model(sorted_beams, viewport_w, apoyos)
        if model_result.get("modelPositions"):
            return model_result

    n = len(sorted_beams or [])
    viewport_w = max(320.0, float(viewport_w or 640.0))
    if not n:
        return {"layouts": [], "contentWidthPx": viewport_w, "needsScroll": False}

    gaps_total = BEAM_GAP_PX * max(n - 1, 0)
    pool_in_viewport = viewport_w - CANVAS_SIDE_PAD_PX * 2.0 - gaps_total
    lengths = [_beam_len_m(b) for b in sorted_beams]
    total_len = sum(lengths)

    if total_len <= 1e-9:
        widths = [BEAM_SLOT_PX] * n
        content_w = CANVAS_SIDE_PAD_PX * 2.0 + gaps_total + sum(widths)
        needs_scroll = content_w > viewport_w + 1.0
    else:
        fit_widths = [(lengths[i] / total_len) * pool_in_viewport for i in range(n)]
        min_fit = min(fit_widths) if fit_widths else 0.0
        if min_fit >= FIT_MIN_SLOT_PX:
            content_w = viewport_w
            widths = fit_widths
            needs_scroll = False
        else:
            widths = [
                max(MIN_BEAM_SLOT_PX, (lengths[i] / total_len) * n * MIN_BEAM_SLOT_PX)
                for i in range(n)
            ]
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
    }


def model_u_span(sorted_beams, apoyos):
    """Rango ``[u_min, u_max]`` proyectado sobre ``view.RightDirection``."""
    values = []
    for beam in sorted_beams or []:
        for key in ("uStart", "uEnd"):
            v = beam.get(key)
            if v is None:
                continue
            try:
                values.append(float(v))
            except (TypeError, ValueError):
                pass
    for ap in apoyos or []:
        v = ap.get("uView")
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


def compute_layout_model(sorted_beams, viewport_w, apoyos):
    """Layout horizontal alineado con posición real en la vista activa."""
    n = len(sorted_beams or [])
    viewport_w = max(320.0, float(viewport_w or 640.0))
    span = model_u_span(sorted_beams, apoyos)
    if not n or span is None:
        return {"modelPositions": False}

    u_min, u_max = span
    u_range = float(u_max) - float(u_min)
    pool_viewport = viewport_w - CANVAS_SIDE_PAD_PX * 2.0
    base_scale = pool_viewport / u_range

    raw_widths = []
    for beam in sorted_beams:
        try:
            u0 = float(beam.get("uStart", u_min))
            u1 = float(beam.get("uEnd", u0))
        except (TypeError, ValueError):
            raw_widths.append(BEAM_SLOT_PX)
            continue
        raw_widths.append(max(0.0, u1 - u0) * base_scale)

    min_w = min(raw_widths) if raw_widths else 0.0
    if min_w >= FIT_MIN_SLOT_PX:
        content_w = viewport_w
        scale = base_scale
        needs_scroll = False
    else:
        if min_w > 1e-9:
            scale = base_scale * (FIT_MIN_SLOT_PX / min_w)
        else:
            scale = base_scale
        content_w = CANVAS_SIDE_PAD_PX * 2.0 + u_range * scale
        needs_scroll = content_w > viewport_w + 1.0

    layouts = _layout_from_model_u(sorted_beams, u_min, u_max, scale, content_w)
    return {
        "layouts": layouts,
        "contentWidthPx": content_w,
        "needsScroll": needs_scroll,
        "modelPositions": True,
        "modelUMin": u_min,
        "modelUMax": u_max,
    }


def tramo_span(layouts, tramo, content_width_px):
    first = layouts[tramo["beamIndices"][0]]
    last = layouts[tramo["beamIndices"][-1]]
    left_pct = first["leftPct"]
    right_pct = last["leftPct"] + last["widthPct"]
    if tramo.get("edgeStart") == "half":
        left_pct = first["leftPct"] + first["widthPct"] * 0.5
    if tramo.get("edgeEnd") == "half":
        right_pct = last["leftPct"] + last["widthPct"] * 0.5
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

    apoyo_by_id = {a["id"]: a for a in (apoyos or []) if a.get("id")}
    chain = []
    last_id = None
    for beam in sorted_beams:
        for key in ("colStart", "colEnd"):
            aid = beam.get(key)
            if not aid or aid == last_id:
                continue
            ap = apoyo_by_id.get(aid)
            u = ap.get("uView") if ap else None
            if u is None:
                continue
            pct = model_u_to_left_pct(u, u_min, u_max, content_w)
            chain.append({"id": aid, "pct": pct})
            last_id = aid
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
