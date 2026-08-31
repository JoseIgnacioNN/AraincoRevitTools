# -*- coding: utf-8 -*-
"""Layout del canvas de elevación — escala unificada (vista Right/Up)."""

from __future__ import division

SECTION_RAIL_PAD_PX = 16.0
SECTION_RAIL_WIDTH_PX = 340.0
SECTION_CTRL_WIDTH_PX = SECTION_RAIL_WIDTH_PX - SECTION_RAIL_PAD_PX

CANVAS_PAD_PX = 28.0
LABEL_BAND_PX = 40.0
ELEVATION_HEAD_PX = 26.0
MIN_MEMBER_PX = 4.0
# Escala mínima (px/pie) y fallback de altura del body
MIN_SCALE_PX_PER_FT = 18.0
FIT_VIEWPORT_H = 520.0

PREVIEW_CANVAS_H = 222.0

# Recorte horizontal de contexto (vigas / losas / fundaciones) en modelo (mm).
# Desde el centroide U del set de columnas → ±mitad a cada lado.
CONTEXT_CLIP_HALF_MM = 2000.0


def column_canvas_label(idx):
    return u"Columna {0}".format(int(idx) + 1)


def _mm_to_ft(mm):
    return float(mm) / 304.8


def representative_column_centroid_u(members, selected_ids=None):
    """
    Centroide horizontal (eje Right / ``u``, pies) del set de columnas.

    Prioriza columnas en ``selected_ids``; si no hay seleccionadas, usa todas.
    """
    cols = [m for m in (members or []) if m.get("kind") == u"column"]
    if not cols:
        return None
    ids = selected_ids or set()
    picked = [m for m in cols if m.get("id") in ids]
    if not picked:
        picked = cols

    # Media de centros, ponderada por área de alzado (spanU × spanV) si está.
    num = 0.0
    den = 0.0
    for m in picked:
        try:
            um = float(m.get("uMid"))
        except Exception:
            try:
                um = 0.5 * (float(m["uMin"]) + float(m["uMax"]))
            except Exception:
                continue
        try:
            w = max(1e-6, abs(float(m.get("spanU_ft") or (float(m["uMax"]) - float(m["uMin"])))))
            h = max(1e-6, abs(float(m.get("spanV_ft") or (float(m["vMax"]) - float(m["vMin"])))))
            weight = w * h
        except Exception:
            weight = 1.0
        num += um * weight
        den += weight
    if den < 1e-12:
        return None
    return num / den


def clip_context_members_horizontal(members, selected_ids=None, half_mm=None):
    """
    Recorta elementos de contexto (no columnas) a ±``half_mm`` del centroide
    horizontal de las columnas (seleccionadas / todas).

    Las columnas se conservan sin recorte.
    Elementos de contexto totalmente fuera de la banda se omiten.
    Devuelve ``(members_clipped, clip_meta)``.
    """
    half_mm = CONTEXT_CLIP_HALF_MM if half_mm is None else float(half_mm)
    half_ft = _mm_to_ft(half_mm)
    members = list(members or [])
    u_c = representative_column_centroid_u(members, selected_ids)
    meta = {
        "applied": False,
        "uCenter": u_c,
        "halfMm": half_mm,
        "uLo": None,
        "uHi": None,
    }
    if u_c is None or half_ft <= 0:
        return members, meta

    u_lo = float(u_c) - half_ft
    u_hi = float(u_c) + half_ft
    meta["applied"] = True
    meta["uLo"] = u_lo
    meta["uHi"] = u_hi

    out = []
    for m in members:
        kind = m.get("kind") or u"column"
        if kind == u"column":
            out.append(m)
            continue
        try:
            umin = float(m["uMin"])
            umax = float(m["uMax"])
        except Exception:
            continue
        if umax < umin:
            umin, umax = umax, umin
        # Intersección con banda de recorte
        c0 = max(umin, u_lo)
        c1 = min(umax, u_hi)
        if c1 - c0 < 1e-6:
            continue
        clipped = dict(m)
        clipped["uMin"] = c0
        clipped["uMax"] = c1
        clipped["uMid"] = 0.5 * (c0 + c1)
        clipped["spanU_ft"] = c1 - c0
        clipped["clippedHorizontal"] = True
        clipped["uMinOrig"] = umin
        clipped["uMaxOrig"] = umax
        out.append(clipped)
    return out, meta


def _members_uv_span(members):
    us0, us1, vs0, vs1 = [], [], [], []
    for m in members or []:
        try:
            us0.append(float(m["uMin"]))
            us1.append(float(m["uMax"]))
            vs0.append(float(m["vMin"]))
            vs1.append(float(m["vMax"]))
        except Exception:
            continue
    if not us0:
        return None
    return min(us0), max(us1), min(vs0), max(vs1)


def compute_elevation_layout(members, viewport_w, viewport_h=None):
    """
    Posiciona cada miembro con **misma escala isométrica** px/pie en U y V
    (proporciones y distancias relativas como en la vista activa).

    Coordenadas de modelo en pies: ``u`` (Right), ``v`` (Up).
    Canvas Y crece hacia abajo ⇒ ``y = top0 + (vMax - v) * scale``.
    """
    viewport_w = max(320.0, float(viewport_w or 640.0))
    viewport_h = max(280.0, float(viewport_h or FIT_VIEWPORT_H))
    members = list(members or [])
    if not members:
        return {
            "layouts": [],
            "contentWidthPx": viewport_w,
            "contentHeightPx": viewport_h,
            "needsScroll": False,
            "scalePxPerFt": 1.0,
            "modelPositions": False,
        }

    span = _members_uv_span(members)
    if span is None:
        return {
            "layouts": [],
            "contentWidthPx": viewport_w,
            "contentHeightPx": viewport_h,
            "needsScroll": False,
            "scalePxPerFt": 1.0,
            "modelPositions": False,
        }

    u_min, u_max, v_min, v_max = span
    range_u = max(u_max - u_min, 1e-6)
    range_v = max(v_max - v_min, 1e-6)

    # Área útil (sin labels inferiores ni cabecera)
    pool_w = max(80.0, viewport_w - CANVAS_PAD_PX * 2.0)
    pool_h = max(80.0, viewport_h - CANVAS_PAD_PX * 2.0 - LABEL_BAND_PX - ELEVATION_HEAD_PX)

    scale_fit = min(pool_w / range_u, pool_h / range_v)
    if scale_fit < MIN_SCALE_PX_PER_FT:
        # No comprimir por debajo del mínimo: activar scroll
        scale = MIN_SCALE_PX_PER_FT
    else:
        scale = scale_fit

    content_w = CANVAS_PAD_PX * 2.0 + range_u * scale
    content_h = (
        ELEVATION_HEAD_PX
        + CANVAS_PAD_PX
        + range_v * scale
        + CANVAS_PAD_PX
        + LABEL_BAND_PX
    )
    # Si el contenido cabe, centrar en el viewport
    if content_w < viewport_w:
        content_w = viewport_w
    needs_scroll = content_w > viewport_w + 1.0 or content_h > viewport_h + 1.0

    # Origen del dibujo (esquina inf-izq del span de modelo en canvas)
    # left = pad + (u - u_min) * scale
    # Alinear horizontalmente el bloque del modelo al centro si sobra ancho
    model_w = range_u * scale
    origin_x = CANVAS_PAD_PX
    if content_w > model_w + CANVAS_PAD_PX * 2.0:
        origin_x = (content_w - model_w) * 0.5

    elev_top = ELEVATION_HEAD_PX + CANVAS_PAD_PX * 0.5
    # top_of_v_max
    top0 = elev_top

    layouts = []
    for idx, m in enumerate(members):
        try:
            um0 = float(m["uMin"])
            um1 = float(m["uMax"])
            vm0 = float(m["vMin"])
            vm1 = float(m["vMax"])
        except Exception:
            continue
        left = origin_x + (um0 - u_min) * scale
        width = max(MIN_MEMBER_PX, (um1 - um0) * scale)
        top = top0 + (v_max - vm1) * scale
        height = max(MIN_MEMBER_PX, (vm1 - vm0) * scale)
        layouts.append({
            "idx": idx,
            "id": m.get("id"),
            "kind": m.get("kind"),
            "leftPx": left,
            "topPx": top,
            "widthPx": width,
            "heightPx": height,
            "centerPx": left + width * 0.5,
            "bottomPx": top + height,
        })

    return {
        "layouts": layouts,
        "contentWidthPx": content_w,
        "contentHeightPx": max(content_h, elev_top + range_v * scale + LABEL_BAND_PX + 16.0),
        "needsScroll": needs_scroll,
        "scalePxPerFt": scale,
        "modelPositions": True,
        "modelUMin": u_min,
        "modelUMax": u_max,
        "modelVMin": v_min,
        "modelVMax": v_max,
        "elevTopPx": elev_top,
        "elevBottomPx": elev_top + range_v * scale,
        "originXPx": origin_x,
    }


def compute_column_slots(columns, viewport_w):
    """Fallback legacy (sin coords de vista) — slots espaciados."""
    n = len(columns or [])
    viewport_w = max(320.0, float(viewport_w or 640.0))
    if not n:
        return {
            "layouts": [],
            "contentWidthPx": viewport_w,
            "needsScroll": False,
            "modelPositions": False,
        }
    gap = 20.0
    pad = 16.0
    slot = max(120.0, min(180.0, (viewport_w - pad * 2.0 - gap * max(n - 1, 0)) / float(n)))
    content_w = pad * 2.0 + gap * max(n - 1, 0) + slot * n
    if content_w < viewport_w:
        content_w = viewport_w
    start = (content_w - (gap * max(n - 1, 0) + slot * n)) * 0.5
    cursor = start
    layouts = []
    for idx, col in enumerate(columns or []):
        layouts.append({
            "idx": idx,
            "id": col.get("id"),
            "kind": col.get("kind") or u"column",
            "leftPx": cursor,
            "widthPx": slot,
            "centerPx": cursor + slot * 0.5,
            "topPx": 40.0,
            "heightPx": 220.0,
            "bottomPx": 260.0,
        })
        cursor += slot + gap
    return {
        "layouts": layouts,
        "contentWidthPx": content_w,
        "contentHeightPx": 320.0,
        "needsScroll": content_w > viewport_w + 1.0,
        "modelPositions": False,
    }
