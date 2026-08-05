# -*- coding: utf-8 -*-
"""
Geometría pura: corte(s) con traslape (sin dependencia de Revit).

Modos de solape:
  - ``symmetric``: ±L/2 en cada extremo del corte (comportamiento histórico).
  - ``endpoint_prev``: todo L estira el tramo anterior (termina en C+L).
  - ``endpoint_next``: todo L estira el tramo siguiente (empieza en C−L).

Usado por ``dividir_rebar_punto_core`` y por validación offline.
"""

from __future__ import print_function

_MM_TO_FT = 1.0 / 304.8
_FT_TO_MM = 304.8

LAP_MODE_SYMMETRIC = u"symmetric"
LAP_MODE_ENDPOINT_PREV = u"endpoint_prev"
LAP_MODE_ENDPOINT_NEXT = u"endpoint_next"
LAP_MODES = (
    LAP_MODE_SYMMETRIC,
    LAP_MODE_ENDPOINT_PREV,
    LAP_MODE_ENDPOINT_NEXT,
)


def normalize_lap_mode(mode):
    """Normaliza a una clave de ``LAP_MODES`` (default simétrico)."""
    try:
        key = u"{0}".format(mode or u"").strip().lower()
    except Exception:
        key = u""
    aliases = {
        u"symmetric": LAP_MODE_SYMMETRIC,
        u"simetrico": LAP_MODE_SYMMETRIC,
        u"simétrico": LAP_MODE_SYMMETRIC,
        u"half": LAP_MODE_SYMMETRIC,
        u"l2": LAP_MODE_SYMMETRIC,
        u"endpoint_prev": LAP_MODE_ENDPOINT_PREV,
        u"prev": LAP_MODE_ENDPOINT_PREV,
        u"anterior": LAP_MODE_ENDPOINT_PREV,
        u"forward": LAP_MODE_ENDPOINT_PREV,
        u"endpoint_next": LAP_MODE_ENDPOINT_NEXT,
        u"next": LAP_MODE_ENDPOINT_NEXT,
        u"siguiente": LAP_MODE_ENDPOINT_NEXT,
        u"backward": LAP_MODE_ENDPOINT_NEXT,
    }
    return aliases.get(key, LAP_MODE_SYMMETRIC)


def lap_zone_around_cut(cut, lap_len, lap_mode=None):
    """
    Intervalo [a, b] del solape en la centerline para un corte.

    Returns:
        (a, b) misma unidad que cut/lap.
    """
    mode = normalize_lap_mode(lap_mode)
    c = float(cut)
    lap = float(lap_len)
    half = 0.5 * lap
    if mode == LAP_MODE_ENDPOINT_PREV:
        return (c, c + lap)
    if mode == LAP_MODE_ENDPOINT_NEXT:
        return (c - lap, c)
    return (c - half, c + half)


def mm_to_internal(mm):
    return float(mm) * _MM_TO_FT


def ft_to_mm(ft):
    return float(ft) * _FT_TO_MM


def validate_cut_with_lap(total_len, cut_dist, lap_len, min_piece=0.0, lap_mode=None):
    """
    Valida corte a ``cut_dist`` sobre longitud ``total_len`` con solape ``lap_len``.

    Todas las longitudes en la misma unidad.

    Returns:
        (ok, mensaje, half, len_a, len_b)
    """
    try:
        L = float(total_len)
        c = float(cut_dist)
        lap = float(lap_len)
        mn = float(min_piece or 0.0)
    except Exception:
        return False, u"Parámetros numéricos inválidos.", 0.0, 0.0, 0.0
    if L <= 0 or lap <= 0:
        return False, u"Longitud de barra o traslape inválidos.", 0.0, 0.0, 0.0
    if c <= 0 or c >= L:
        return (
            False,
            u"El punto de corte debe quedar estrictamente entre los extremos.",
            0.0,
            0.0,
            0.0,
        )
    mode = normalize_lap_mode(lap_mode)
    half = 0.5 * lap
    if mode == LAP_MODE_ENDPOINT_PREV:
        if (c + lap) > L + 1e-9:
            return (
                False,
                u"Corte demasiado cerca del final (hace falta ≥ L libre tras el corte).",
                half,
                0.0,
                0.0,
            )
        len_a = c + lap
        len_b = L - c
    elif mode == LAP_MODE_ENDPOINT_NEXT:
        if c < lap - 1e-9:
            return (
                False,
                u"Corte demasiado cerca del inicio (hace falta ≥ L libre antes del corte).",
                half,
                0.0,
                0.0,
            )
        len_a = c
        len_b = (L - c) + lap
    else:
        if c < half:
            return (
                False,
                u"Corte demasiado cerca del inicio (hace falta ≥ lap/2 libre).",
                half,
                0.0,
                0.0,
            )
        if (L - c) < half:
            return (
                False,
                u"Corte demasiado cerca del final (hace falta ≥ lap/2 libre).",
                half,
                0.0,
                0.0,
            )
        len_a = c + half
        len_b = (L - c) + half
    if len_a < mn or len_b < mn:
        return (
            False,
            u"Tras el traslape, algún tramo quedaría más corto que el mínimo permitido.",
            half,
            len_a,
            len_b,
        )
    return True, u"", half, len_a, len_b


def split_distances_with_lap(total_len, cut_dist, lap_len, lap_mode=None):
    """
    Intervalos [start, end] para tramo A y B según ``lap_mode``.
    """
    ok, msg, half, _la, _lb = validate_cut_with_lap(
        total_len, cut_dist, lap_len, 0.0, lap_mode=lap_mode
    )
    if not ok:
        raise ValueError(msg)
    intervals = piece_intervals_with_lap(
        total_len, [cut_dist], lap_len, lap_mode=lap_mode
    )
    if len(intervals) < 2:
        raise ValueError(u"No se generaron tramos.")
    return intervals[0], intervals[1], half


def overlap_length(interval_a, interval_b):
    """Longitud de solape entre dos intervalos [a0,a1] y [b0,b1]."""
    a0, a1 = float(interval_a[0]), float(interval_a[1])
    b0, b1 = float(interval_b[0]), float(interval_b[1])
    lo = max(a0, b0)
    hi = min(a1, b1)
    return max(0.0, hi - lo)


def _sorted_unique_cuts(cuts, tol=1e-9):
    vals = []
    for c in cuts or []:
        try:
            v = float(c)
        except Exception:
            continue
        vals.append(v)
    vals.sort()
    out = []
    for v in vals:
        if not out or abs(v - out[-1]) > tol:
            out.append(v)
    return out


def validate_cuts_with_lap(total_len, cuts, lap_len, min_piece=0.0, lap_mode=None):
    """
    Valida una lista de cortes (misma unidad que ``total_len`` / ``lap_len``).

    Returns:
        (ok, mensaje, cuts_sorted)
    """
    try:
        L = float(total_len)
        lap = float(lap_len)
        mn = float(min_piece or 0.0)
    except Exception:
        return False, u"Parámetros numéricos inválidos.", []
    if L <= 0 or lap <= 0:
        return False, u"Longitud de barra o traslape inválidos.", []
    sorted_cuts = _sorted_unique_cuts(cuts)
    if not sorted_cuts:
        return False, u"Indique al menos un punto de corte.", []
    mode = normalize_lap_mode(lap_mode)
    half = 0.5 * lap
    for i, c in enumerate(sorted_cuts):
        if c <= 0 or c >= L:
            return (
                False,
                u"El corte {0} debe quedar entre los extremos.".format(i + 1),
                sorted_cuts,
            )
        if mode == LAP_MODE_ENDPOINT_PREV:
            if (c + lap) > L + 1e-9:
                return (
                    False,
                    u"Corte {0} demasiado cerca del final (≥ L tras el corte).".format(
                        i + 1
                    ),
                    sorted_cuts,
                )
        elif mode == LAP_MODE_ENDPOINT_NEXT:
            if c < lap - 1e-9:
                return (
                    False,
                    u"Corte {0} demasiado cerca del inicio (≥ L antes del corte).".format(
                        i + 1
                    ),
                    sorted_cuts,
                )
        else:
            if c < half:
                return (
                    False,
                    u"Corte {0} demasiado cerca del inicio (≥ lap/2).".format(i + 1),
                    sorted_cuts,
                )
            if (L - c) < half:
                return (
                    False,
                    u"Corte {0} demasiado cerca del final (≥ lap/2).".format(i + 1),
                    sorted_cuts,
                )
        if i > 0 and (c - sorted_cuts[i - 1]) < lap:
            return (
                False,
                u"Separación entre cortes {0} y {1} < traslape.".format(i, i + 1),
                sorted_cuts,
            )
    prev = 0.0
    for i, c in enumerate(sorted_cuts + [L]):
        span = c - prev
        if span < mn:
            return (
                False,
                u"Vano {0} más corto que el mínimo ({1}).".format(i + 1, mn),
                sorted_cuts,
            )
        prev = c
    return True, u"", sorted_cuts


def build_spans_mm(total_mm, cuts_mm):
    """
    Vanos nominales entre 0 / cortes / L.

    Returns:
        list of dict: label (T1…), start_mm, end_mm, length_mm, closing_cut_index (-1 = fin)
    """
    L = float(total_mm)
    sorted_cuts = _sorted_unique_cuts(cuts_mm)
    spans = []
    prev = 0.0
    for i, c in enumerate(sorted_cuts):
        spans.append(
            {
                u"label": tramo_label(i),
                u"start_mm": prev,
                u"end_mm": c,
                u"length_mm": c - prev,
                u"closing_cut_index": i,
            }
        )
        prev = c
    spans.append(
        {
            u"label": tramo_label(len(sorted_cuts)),
            u"start_mm": prev,
            u"end_mm": L,
            u"length_mm": L - prev,
            u"closing_cut_index": -1,
        }
    )
    return spans


def tramo_label(span_index):
    """Etiqueta de tramo alineada UI ↔ canvas (T1, T2, …)."""
    try:
        i = int(span_index)
    except Exception:
        i = 0
    return u"T{0}".format(max(0, i) + 1)


def span_midpoint_mm(span):
    """Punto medio del vano en mm (para etiqueta en canvas)."""
    if not span:
        return 0.0
    try:
        a = float(span.get(u"start_mm") or 0.0)
        b = float(span.get(u"end_mm") or 0.0)
    except Exception:
        return 0.0
    return 0.5 * (a + b)


def span_index_at_mm(total_mm, cuts_mm, mm):
    """
    Índice de tramo (0-based) que contiene ``mm`` a lo largo de la centerline.

    En un corte exacto pertenece al tramo siguiente (el que empieza ahí),
    salvo el final L que queda en el último tramo.
    """
    spans = build_spans_mm(total_mm, cuts_mm)
    if not spans:
        return -1
    try:
        x = float(mm)
        L = float(total_mm)
    except Exception:
        return -1
    if L <= 0:
        return -1
    if x <= 0:
        return 0
    if x >= L:
        return len(spans) - 1
    for i, sp in enumerate(spans):
        a = float(sp.get(u"start_mm") or 0.0)
        b = float(sp.get(u"end_mm") or 0.0)
        if i < len(spans) - 1:
            if a <= x < b:
                return i
        else:
            if a <= x <= b:
                return i
    return len(spans) - 1


def piece_intervals_with_lap(total_len, cuts, lap_len, lap_mode=None):
    """
    Intervalos [start, end] de cada tramo fabricado según ``lap_mode``.

    Returns:
        list of (d0, d1)
    """
    mode = normalize_lap_mode(lap_mode)
    ok, msg, sorted_cuts = validate_cuts_with_lap(
        total_len, cuts, lap_len, 0.0, lap_mode=mode
    )
    if not ok:
        raise ValueError(msg)
    L = float(total_len)
    lap = float(lap_len)
    half = 0.5 * lap
    n = len(sorted_cuts)
    intervals = []
    for i in range(n + 1):
        if mode == LAP_MODE_ENDPOINT_PREV:
            if i == 0:
                d0 = 0.0
                d1 = L if n == 0 else min(L, sorted_cuts[0] + lap)
            elif i == n:
                d0 = sorted_cuts[n - 1]
                d1 = L
            else:
                d0 = sorted_cuts[i - 1]
                d1 = min(L, sorted_cuts[i] + lap)
        elif mode == LAP_MODE_ENDPOINT_NEXT:
            if i == 0:
                d0 = 0.0
                d1 = L if n == 0 else sorted_cuts[0]
            elif i == n:
                d0 = max(0.0, sorted_cuts[n - 1] - lap)
                d1 = L
            else:
                d0 = max(0.0, sorted_cuts[i - 1] - lap)
                d1 = sorted_cuts[i]
        else:
            if i == 0:
                d0 = 0.0
                d1 = L if n == 0 else min(L, sorted_cuts[0] + half)
            elif i == n:
                d0 = max(0.0, sorted_cuts[n - 1] - half)
                d1 = L
            else:
                d0 = max(0.0, sorted_cuts[i - 1] - half)
                d1 = min(L, sorted_cuts[i] + half)
        intervals.append((d0, d1))
    return intervals


def set_span_length_mm(
    total_mm,
    cuts_mm,
    span_index,
    new_len_mm,
    lap_mm,
    min_piece_mm=100.0,
    lap_mode=None,
):
    """
    Ajusta el largo del vano ``span_index`` moviendo el corte límite.
    El vano vecino absorbe el delta. Suma de vanos = total.

    Returns:
        (ok, mensaje, cuts_mm_nuevos)
    """
    mode = normalize_lap_mode(lap_mode)
    try:
        L = float(total_mm)
        lap = float(lap_mm)
        mn = float(min_piece_mm or 0.0)
        target = float(new_len_mm)
    except Exception:
        return False, u"Parámetros numéricos inválidos.", list(cuts_mm or [])
    sorted_cuts = _sorted_unique_cuts(cuts_mm)
    if not sorted_cuts:
        return False, u"Agregue al menos un corte.", []
    spans = build_spans_mm(L, sorted_cuts)
    if span_index < 0 or span_index >= len(spans):
        return False, u"Vano inválido.", sorted_cuts
    current = float(spans[span_index][u"length_mm"])
    delta = target - current
    if abs(delta) < 1e-9:
        return True, u"", sorted_cuts
    next_cuts = list(sorted_cuts)
    if span_index < len(spans) - 1:
        cut_idx = span_index
        new_cut = next_cuts[cut_idx] + delta
        prev_b = 0.0 if cut_idx == 0 else next_cuts[cut_idx - 1]
        next_b = L if cut_idx + 1 >= len(next_cuts) else next_cuts[cut_idx + 1]
        if new_cut - prev_b < mn:
            return (
                False,
                u"Vano demasiado corto (mín. {0:.0f} mm).".format(mn),
                sorted_cuts,
            )
        if next_b - new_cut < mn:
            return False, u"El vano vecino quedaría bajo el mínimo.", sorted_cuts
        next_cuts[cut_idx] = new_cut
    else:
        cut_idx = len(next_cuts) - 1
        new_cut = next_cuts[cut_idx] - delta
        prev_b = 0.0 if cut_idx == 0 else next_cuts[cut_idx - 1]
        if new_cut - prev_b < mn:
            return False, u"El vano anterior quedaría bajo el mínimo.", sorted_cuts
        if L - new_cut < mn:
            return (
                False,
                u"Vano demasiado corto (mín. {0:.0f} mm).".format(mn),
                sorted_cuts,
            )
        next_cuts[cut_idx] = new_cut
    ok, msg, final = validate_cuts_with_lap(L, next_cuts, lap, mn, lap_mode=mode)
    if not ok:
        return False, msg, sorted_cuts
    return True, u"", final


# ---------------------------------------------------------------------------
# Esquema 2D de la centerline (proyección ortogonal + longitud de arco)
# ---------------------------------------------------------------------------


def _dist2(a, b):
    return ((float(a[0]) - float(b[0])) ** 2 + (float(a[1]) - float(b[1])) ** 2) ** 0.5


def _dist3(a, b):
    return (
        (float(a[0]) - float(b[0])) ** 2
        + (float(a[1]) - float(b[1])) ** 2
        + (float(a[2]) - float(b[2])) ** 2
    ) ** 0.5


def _v3_dot(a, b):
    return float(a[0]) * float(b[0]) + float(a[1]) * float(b[1]) + float(a[2]) * float(b[2])


def _v3_cross(a, b):
    return (
        float(a[1]) * float(b[2]) - float(a[2]) * float(b[1]),
        float(a[2]) * float(b[0]) - float(a[0]) * float(b[2]),
        float(a[0]) * float(b[1]) - float(a[1]) * float(b[0]),
    )


def _v3_len(a):
    return (_v3_dot(a, a)) ** 0.5


def _v3_norm(a):
    L = _v3_len(a)
    if L < 1e-12:
        return (0.0, 0.0, 0.0)
    return (float(a[0]) / L, float(a[1]) / L, float(a[2]) / L)


def _v3_scale(a, s):
    return (float(a[0]) * s, float(a[1]) * s, float(a[2]) * s)


def orthonormal_frame_from_normal(normal):
    """
    Base U,V en el plano de la barra. V se alinea con +Z mundo cuando es posible
    (arriba del modelo → arriba del esquema).
    """
    n = _v3_norm(normal)
    if _v3_len(n) < 1e-12:
        n = (0.0, 0.0, 1.0)
    world_up = (0.0, 0.0, 1.0)
    if abs(_v3_dot(n, world_up)) > 0.95:
        ref = (0.0, 1.0, 0.0)
    else:
        ref = world_up
    u = _v3_norm(_v3_cross(n, ref))
    if _v3_len(u) < 1e-12:
        u = _v3_norm(_v3_cross(n, (1.0, 0.0, 0.0)))
    v = _v3_norm(_v3_cross(u, n))
    if _v3_dot(v, world_up) < 0.0:
        v = _v3_scale(v, -1.0)
        u = _v3_scale(u, -1.0)
    return u, v, n


def project_points_onto_normal_plane(points_xyz, normal):
    """Proyecta puntos 3D al plano de la barra (normal Revit)."""
    if not points_xyz:
        return [], u"plane"
    u, v, _n = orthonormal_frame_from_normal(normal)
    origin = (
        float(points_xyz[0][0]),
        float(points_xyz[0][1]),
        float(points_xyz[0][2]),
    )
    uv = []
    for p in points_xyz:
        d = (
            float(p[0]) - origin[0],
            float(p[1]) - origin[1],
            float(p[2]) - origin[2],
        )
        uv.append((_v3_dot(d, u), _v3_dot(d, v)))
    return uv, u"normal"


def project_points_to_plan_uv(points_xyz):
    """
    Proyecta puntos 3D a 2D descartando el eje de menor extensión.

    Returns:
        (points_uv, plane_name)  plane_name in {'xy','xz','yz'}
    """
    if not points_xyz:
        return [], u"xy"
    xs = [float(p[0]) for p in points_xyz]
    ys = [float(p[1]) for p in points_xyz]
    zs = [float(p[2]) for p in points_xyz]
    ex = max(xs) - min(xs)
    ey = max(ys) - min(ys)
    ez = max(zs) - min(zs)
    if ez <= ex and ez <= ey:
        plane = u"xy"
        uv = [(x, y) for x, y, _z in zip(xs, ys, zs)]
    elif ey <= ex and ey <= ez:
        plane = u"xz"
        uv = [(x, z) for x, _y, z in zip(xs, ys, zs)]
    else:
        plane = u"yz"
        uv = [(y, z) for _x, y, z in zip(xs, ys, zs)]
    return uv, plane


def polyline_arc_lengths(points_uv):
    """Longitudes de arco acumuladas en el mismo sistema que points_uv."""
    if not points_uv:
        return []
    s = [0.0]
    acc = 0.0
    for i in range(1, len(points_uv)):
        acc += _dist2(points_uv[i - 1], points_uv[i])
        s.append(acc)
    return s


def build_plan_polyline_mm(points_xyz_mm, normal=None):
    """
    Construye polilínea 2D (mm) para el esquema de UI.

    Si ``normal`` está definido, proyecta sobre el plano de la barra (orientación
    fiel al modelo). Si no, cae al plano ortogonal de mayor extensión.

    Returns:
        dict: points_uv [[u,v],...], plane, total_mm, arc_mm [...], flip_v
        o None si no hay puntos suficientes.
    """
    cleaned = []
    for p in points_xyz_mm or []:
        try:
            pt = (float(p[0]), float(p[1]), float(p[2]))
        except Exception:
            continue
        if cleaned and _dist3(cleaned[-1], pt) < 1e-6:
            continue
        cleaned.append(pt)
    if len(cleaned) < 2:
        return None
    flip_v = True
    if normal is not None:
        try:
            n = (float(normal[0]), float(normal[1]), float(normal[2]))
        except Exception:
            n = None
        if n is not None and _v3_len(n) > 1e-12:
            uv, plane = project_points_onto_normal_plane(cleaned, n)
        else:
            uv, plane = project_points_to_plan_uv(cleaned)
            flip_v = plane in (u"xz", u"yz")
    else:
        uv, plane = project_points_to_plan_uv(cleaned)
        flip_v = plane in (u"xz", u"yz")
    if len(uv) >= 2 and polyline_arc_lengths(uv)[-1] < 1e-6:
        uv = [(float(p[0]), float(p[1])) for p in cleaned]
        plane = u"xy"
        flip_v = False
    arcs = polyline_arc_lengths(uv)
    return {
        u"points_uv": [[float(u), float(v)] for u, v in uv],
        u"plane": plane,
        u"total_mm": float(arcs[-1]) if arcs else 0.0,
        u"arc_mm": [float(a) for a in arcs],
        u"flip_v": bool(flip_v),
    }


def fit_polyline_to_canvas(
    points_uv, canvas_w, canvas_h, margin, swap_uv=False, flip_v=False
):
    """
    Escala uniforme + centrado. Devuelve puntos en píxeles canvas y el scale (px/mm).

    ``flip_v``: invierte V al mapear a Y de pantalla (útil si V es Z de Revit).

    Returns:
        (points_px, scale_px_per_mm)
    """
    if not points_uv:
        return [], 1.0
    pts = []
    for p in points_uv:
        u, v = float(p[0]), float(p[1])
        if swap_uv:
            u, v = v, u
        pts.append((u, v))
    us = [p[0] for p in pts]
    vs = [p[1] for p in pts]
    u0, u1 = min(us), max(us)
    v0, v1 = min(vs), max(vs)
    du = max(u1 - u0, 1e-6)
    dv = max(v1 - v0, 1e-6)
    usable_w = max(1.0, float(canvas_w) - 2.0 * float(margin))
    usable_h = max(1.0, float(canvas_h) - 2.0 * float(margin))
    scale = min(usable_w / du, usable_h / dv)
    ox = float(margin) + 0.5 * (usable_w - du * scale) - u0 * scale
    if flip_v:
        # Mayor V → menor Y canvas (arriba en pantalla)
        oy = float(margin) + 0.5 * (usable_h - dv * scale) + v1 * scale
        out = [(ox + p[0] * scale, oy - p[1] * scale) for p in pts]
    else:
        oy = float(margin) + 0.5 * (usable_h - dv * scale) - v0 * scale
        out = [(ox + p[0] * scale, oy + p[1] * scale) for p in pts]
    return out, scale


def point_at_arc_length_uv(points_uv, arc_mm, s_mm):
    """
    Punto (u,v) a longitud de arco s_mm sobre la polilínea.
    ``arc_mm`` = polyline_arc_lengths(points_uv) (opcional; se recalcula si falta).
    """
    if not points_uv:
        return None
    arcs = arc_mm if arc_mm is not None else polyline_arc_lengths(points_uv)
    if not arcs:
        return None
    target = max(0.0, min(float(s_mm), float(arcs[-1])))
    if target <= 1e-12:
        return (float(points_uv[0][0]), float(points_uv[0][1]))
    for i in range(1, len(points_uv)):
        if arcs[i] + 1e-12 >= target:
            seg = arcs[i] - arcs[i - 1]
            if seg < 1e-12:
                return (float(points_uv[i][0]), float(points_uv[i][1]))
            t = (target - arcs[i - 1]) / seg
            u0, v0 = float(points_uv[i - 1][0]), float(points_uv[i - 1][1])
            u1, v1 = float(points_uv[i][0]), float(points_uv[i][1])
            return (u0 + t * (u1 - u0), v0 + t * (v1 - v0))
    return (float(points_uv[-1][0]), float(points_uv[-1][1]))


def tangent_at_arc_length_uv(points_uv, arc_mm, s_mm):
    """Vector tangente unitario (du,dv) en s_mm."""
    if not points_uv or len(points_uv) < 2:
        return (1.0, 0.0)
    arcs = arc_mm if arc_mm is not None else polyline_arc_lengths(points_uv)
    target = max(0.0, min(float(s_mm), float(arcs[-1])))
    for i in range(1, len(points_uv)):
        if arcs[i] + 1e-12 >= target or i == len(points_uv) - 1:
            du = float(points_uv[i][0]) - float(points_uv[i - 1][0])
            dv = float(points_uv[i][1]) - float(points_uv[i - 1][1])
            leng = (du * du + dv * dv) ** 0.5
            if leng < 1e-12:
                continue
            return (du / leng, dv / leng)
    return (1.0, 0.0)


def nearest_arc_length_px(points_px, arc_mm, x, y):
    """
    Proyecta (x,y) canvas sobre la polilínea en px; devuelve (s_mm, dist_px, px, py).
    ``arc_mm`` debe corresponder a los mismos vértices que ``points_px``.
    """
    if not points_px or len(points_px) < 2:
        return 0.0, 1e9, float(x), float(y)
    arcs = arc_mm if arc_mm is not None else polyline_arc_lengths(points_px)
    # Si arcs viene en mm y points en px, las longitudes de segmento deben
    # mapearse por índice (proporción de cada tramo respecto al total mm).
    best_d2 = None
    best_s = 0.0
    best_pt = (float(points_px[0][0]), float(points_px[0][1]))
    total_mm = float(arcs[-1]) if arcs else 0.0
    for i in range(1, len(points_px)):
        x0, y0 = float(points_px[i - 1][0]), float(points_px[i - 1][1])
        x1, y1 = float(points_px[i][0]), float(points_px[i][1])
        dx, dy = x1 - x0, y1 - y0
        seg2 = dx * dx + dy * dy
        if seg2 < 1e-12:
            t = 0.0
            qx, qy = x0, y0
        else:
            t = ((float(x) - x0) * dx + (float(y) - y0) * dy) / seg2
            if t < 0.0:
                t = 0.0
            elif t > 1.0:
                t = 1.0
            qx = x0 + t * dx
            qy = y0 + t * dy
        d2 = (float(x) - qx) ** 2 + (float(y) - qy) ** 2
        s0 = float(arcs[i - 1]) if i - 1 < len(arcs) else 0.0
        s1 = float(arcs[i]) if i < len(arcs) else total_mm
        s = s0 + t * (s1 - s0)
        if best_d2 is None or d2 < best_d2:
            best_d2 = d2
            best_s = s
            best_pt = (qx, qy)
    return best_s, (best_d2 ** 0.5) if best_d2 is not None else 1e9, best_pt[0], best_pt[1]


def subpath_between_arc(points_uv, arc_mm, s0, s1, samples_per_seg=1):
    """
    Vértices de la polilínea entre s0 y s1 (incluye extremos interpolados).
    """
    if not points_uv:
        return []
    a0 = min(float(s0), float(s1))
    a1 = max(float(s0), float(s1))
    arcs = arc_mm if arc_mm is not None else polyline_arc_lengths(points_uv)
    out = []
    p_start = point_at_arc_length_uv(points_uv, arcs, a0)
    if p_start is not None:
        out.append(p_start)
    for i in range(1, len(points_uv) - 1):
        if arcs[i] > a0 + 1e-9 and arcs[i] < a1 - 1e-9:
            out.append((float(points_uv[i][0]), float(points_uv[i][1])))
    p_end = point_at_arc_length_uv(points_uv, arcs, a1)
    if p_end is not None:
        if not out or _dist2(out[-1], p_end) > 1e-9:
            out.append(p_end)
    return out
