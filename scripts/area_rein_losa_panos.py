# -*- coding: utf-8 -*-
"""
Detección de paños (polígonos) para Area Rein. losa.

Parte el contorno exterior (mm, plano Sketch) con segmentos divisores
(ejes de muros/vigas) y devuelve polígonos candidatos a AreaReinforcement.
IronPython 2.7 / Revit — sin dependencias externas.
"""

from __future__ import print_function

import math

_EPS = 1e-6
_EPS_MM = 0.5  # mm


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def shoelace_area_m2(pts):
    """Área absoluta en m² (pts en mm)."""
    if not pts or len(pts) < 3:
        return 0.0
    a = 0.0
    n = len(pts)
    for i in range(n):
        x0, y0 = float(pts[i][0]), float(pts[i][1])
        x1, y1 = float(pts[(i + 1) % n][0]), float(pts[(i + 1) % n][1])
        a += x0 * y1 - x1 * y0
    return abs(a) * 0.5 / 1.0e6


def ensure_ccw(pts):
    if not pts or len(pts) < 3:
        return list(pts or [])
    a = 0.0
    n = len(pts)
    for i in range(n):
        x0, y0 = float(pts[i][0]), float(pts[i][1])
        x1, y1 = float(pts[(i + 1) % n][0]), float(pts[(i + 1) % n][1])
        a += x0 * y1 - x1 * y0
    out = list(pts)
    if a < 0:
        out.reverse()
    return out


def _cross(ax, ay, bx, by):
    return ax * by - ay * bx


def _dot(ax, ay, bx, by):
    return ax * bx + ay * by


def _dist(a, b):
    return math.hypot(float(b[0]) - float(a[0]), float(b[1]) - float(a[1]))


def _line_seg_intersection(a, b, c, d):
    """
    Intersección segmento ab con segmento cd.
    Devuelve (punto, t_ab, u_cd) o None.
    """
    ax, ay = float(a[0]), float(a[1])
    bx, by = float(b[0]), float(b[1])
    cx, cy = float(c[0]), float(c[1])
    dx, dy = float(d[0]), float(d[1])
    r_x, r_y = bx - ax, by - ay
    s_x, s_y = dx - cx, dy - cy
    den = _cross(r_x, r_y, s_x, s_y)
    if abs(den) < _EPS:
        return None
    qp_x, qp_y = cx - ax, cy - ay
    t = _cross(qp_x, qp_y, s_x, s_y) / den
    u = _cross(qp_x, qp_y, r_x, r_y) / den
    if t < -_EPS or t > 1.0 + _EPS or u < -_EPS or u > 1.0 + _EPS:
        return None
    t = max(0.0, min(1.0, t))
    return (ax + t * r_x, ay + t * r_y), t, u


def _line_edge_intersection(a, b, c, d):
    """
    Intersección de la recta infinita ab con el segmento cd.
    Devuelve (punto, t_line, u_edge) o None. t_line puede ser cualquiera.
    """
    ax, ay = float(a[0]), float(a[1])
    bx, by = float(b[0]), float(b[1])
    cx, cy = float(c[0]), float(c[1])
    dx, dy = float(d[0]), float(d[1])
    r_x, r_y = bx - ax, by - ay
    s_x, s_y = dx - cx, dy - cy
    den = _cross(r_x, r_y, s_x, s_y)
    if abs(den) < _EPS:
        return None
    qp_x, qp_y = cx - ax, cy - ay
    t = _cross(qp_x, qp_y, s_x, s_y) / den
    u = _cross(qp_x, qp_y, r_x, r_y) / den
    if u < -_EPS or u > 1.0 + _EPS:
        return None
    u = max(0.0, min(1.0, u))
    return (ax + t * r_x, ay + t * r_y), t, u


def _point_on_poly_boundary(pt, poly, tol=_EPS_MM):
    n = len(poly)
    for i in range(n):
        a = poly[i]
        b = poly[(i + 1) % n]
        if _dist(pt, a) <= tol or _dist(pt, b) <= tol:
            return True
        # proyección en segmento
        ax, ay = float(a[0]), float(a[1])
        bx, by = float(b[0]), float(b[1])
        px, py = float(pt[0]), float(pt[1])
        abx, aby = bx - ax, by - ay
        ab2 = abx * abx + aby * aby
        if ab2 < _EPS:
            continue
        t = ((px - ax) * abx + (py - ay) * aby) / ab2
        if t < 0.0 or t > 1.0:
            continue
        cx, cy = ax + t * abx, ay + t * aby
        if math.hypot(px - cx, py - cy) <= tol:
            return True
    return False


def _dedupe_hits(hits, tol=_EPS_MM):
    """hits: list of (pt, t_line, edge_i, u_edge)."""
    if not hits:
        return []
    hits = sorted(hits, key=lambda h: h[1])
    out = [hits[0]]
    for h in hits[1:]:
        if _dist(h[0], out[-1][0]) > tol:
            out.append(h)
    return out


def _insert_points_on_ring(poly, pts_with_edge):
    """
    Inserta puntos en el anillo. pts_with_edge: [(pt, edge_i, u), ...]
    Devuelve nuevo anillo (lista de pts) y mapa pt_key -> index.
    """
    n = len(poly)
    by_edge = {}
    for pt, ei, u in pts_with_edge:
        by_edge.setdefault(int(ei), []).append((float(u), pt))
    new_ring = []
    for i in range(n):
        new_ring.append(poly[i])
        extras = by_edge.get(i, [])
        extras.sort(key=lambda x: x[0])
        for _u, pt in extras:
            if _dist(pt, poly[i]) <= _EPS_MM:
                continue
            if _dist(pt, poly[(i + 1) % n]) <= _EPS_MM:
                continue
            if new_ring and _dist(pt, new_ring[-1]) <= _EPS_MM:
                continue
            new_ring.append((float(pt[0]), float(pt[1])))
    # quitar duplicado cierre
    if len(new_ring) > 1 and _dist(new_ring[0], new_ring[-1]) <= _EPS_MM:
        new_ring = new_ring[:-1]
    return new_ring


def _find_index(ring, pt, tol=_EPS_MM):
    for i, p in enumerate(ring):
        if _dist(p, pt) <= tol:
            return i
    return -1


def split_polygon_by_line(poly, a, b):
    """
    Parte un polígono con la recta que pasa por a-b.
    Con 2 intersecciones → 2 polígonos. Si no corta → [poly].
    """
    if not poly or len(poly) < 3:
        return [poly]
    ax, ay = float(a[0]), float(a[1])
    bx, by = float(b[0]), float(b[1])
    if math.hypot(bx - ax, by - ay) < _EPS:
        return [list(poly)]

    return _split_once_two_hits(list(poly), None, None, a, b)


def _split_once_two_hits(poly, _h0, _h1, a, b):
    """Parte poly buscando 2 intersecciones de la recta a-b con el contorno."""
    if not poly or len(poly) < 3:
        return [poly]

    hits = []
    n = len(poly)
    for i in range(n):
        c = poly[i]
        d = poly[(i + 1) % n]
        res = _line_edge_intersection(a, b, c, d)
        if res is None:
            continue
        pt, t, u = res
        if u > 1.0 - 1e-8:
            continue
        hits.append((pt, t, i, u))
    hits = _dedupe_hits(hits)
    if len(hits) < 2:
        return [list(poly)]

    # Elegir el mejor par (cuerda que cruza cerca del segmento divisor)
    h0, h1 = hits[0], hits[1]
    if len(hits) > 2:
        best = None
        best_score = 1e18
        for i in range(len(hits)):
            for j in range(i + 1, len(hits)):
                ti, tj = hits[i][1], hits[j][1]
                pen = 0.0
                for tt in (ti, tj):
                    if tt < -0.05:
                        pen += abs(tt)
                    if tt > 1.05:
                        pen += abs(tt - 1.0)
                score = abs(ti - tj) + pen * 1000.0
                if score < best_score:
                    best_score = score
                    best = (hits[i], hits[j])
        if best is not None:
            h0, h1 = best
    if h0[1] > h1[1]:
        h0, h1 = h1, h0

    pt0, e0, u0 = h0[0], h0[2], h0[3]
    pt1, e1, u1 = h1[0], h1[2], h1[3]

    ring = _insert_points_on_ring(poly, [(pt0, e0, u0), (pt1, e1, u1)])
    i0 = _find_index(ring, pt0)
    i1 = _find_index(ring, pt1)
    if i0 < 0 or i1 < 0 or i0 == i1:
        return [list(poly)]

    def walk(i_from, i_to):
        path = [ring[i_from]]
        n_r = len(ring)
        i = i_from
        guard = 0
        while i != i_to and guard < n_r + 2:
            i = (i + 1) % n_r
            path.append(ring[i])
            guard += 1
        return path

    poly_a = _clean_ring(walk(i0, i1))
    poly_b = _clean_ring(walk(i1, i0))
    out = []
    for p in (poly_a, poly_b):
        if p and len(p) >= 3 and shoelace_area_m2(p) > 1e-8:
            out.append(p)
    return out if out else [list(poly)]


def _clean_ring(pts):
    if not pts:
        return []
    out = []
    for p in pts:
        q = (float(p[0]), float(p[1]))
        if out and _dist(out[-1], q) <= _EPS_MM:
            continue
        out.append(q)
    if len(out) > 1 and _dist(out[0], out[-1]) <= _EPS_MM:
        out = out[:-1]
    return out


def _remove_collinear_mm(pts, tol=_EPS_MM):
    """Elimina vértices colineales (limpieza de anillo; AABB/luz menor no depende)."""
    if not pts or len(pts) < 3:
        return list(pts or [])
    t = float(tol)
    n = len(pts)
    out = []
    for i in range(n):
        prev = pts[(i - 1) % n]
        cur = pts[i]
        nxt = pts[(i + 1) % n]
        ax = float(cur[0]) - float(prev[0])
        ay = float(cur[1]) - float(prev[1])
        bx = float(nxt[0]) - float(cur[0])
        by = float(nxt[1]) - float(cur[1])
        # colineales si cross ≈ 0 y no hay retroceso degenerado
        if abs(_cross(ax, ay, bx, by)) <= t * max(1.0, math.hypot(ax, ay), math.hypot(bx, by)):
            continue
        out.append((float(cur[0]), float(cur[1])))
    if len(out) < 3:
        return _clean_ring(pts)
    return _clean_ring(out)


def split_polygons_by_segment(polygons, a, b):
    """Aplica un corte (recta por a-b) a una lista de polígonos."""
    out = []
    for poly in polygons or []:
        parts = split_polygon_by_line(poly, a, b)
        out.extend(parts)
    return out


def detectar_panos_mm(outer_pts_mm, divider_segs, min_area_m2=1.0):
    """
    outer_pts_mm: lista (x,y) mm del perímetro exterior (cerrado sin repetir).
    divider_segs: lista de dict con keys a, b (pt mm), kind, width_mm, eid.
    Devuelve lista de dict: id, label, pts, area_m2, kind_divisors...
    """
    if not outer_pts_mm or len(outer_pts_mm) < 3:
        return []
    outer = ensure_ccw([(float(p[0]), float(p[1])) for p in outer_pts_mm])
    polys = [outer]
    for seg in divider_segs or []:
        a = seg.get(u"a")
        b = seg.get(u"b")
        if a is None or b is None:
            continue
        if _dist(a, b) < 1.0:
            continue
        polys = split_polygons_by_segment(polys, a, b)

    panos = []
    idx = 1
    for poly in polys:
        poly = ensure_ccw(_clean_ring(poly))
        if len(poly) < 3:
            continue
        area = shoelace_area_m2(poly)
        if area + 1e-9 < float(min_area_m2):
            continue
        panos.append(
            {
                u"id": u"P{}".format(idx),
                u"label": u"Paño {}".format(idx),
                u"pts": poly,
                u"area_m2": area,
            }
        )
        idx += 1
    # Orden estable: por centroide X luego Y
    def _key(p):
        pts = p[u"pts"]
        cx = sum(q[0] for q in pts) / float(len(pts))
        cy = sum(q[1] for q in pts) / float(len(pts))
        return (cx, cy)

    panos.sort(key=_key)
    for i, p in enumerate(panos):
        p[u"id"] = u"P{}".format(i + 1)
        p[u"label"] = u"Paño {}".format(i + 1)
    return panos


def inset_polygon_mm(pts, dist_mm):
    """
    Offset uniforme hacia adentro (aprox. por normales de arista).

    Si ``dist≈0``, devuelve una copia de ``pts``.
    Si falla o el polígono colapsa, devuelve ``None`` (no el original).
    """
    if not pts or len(pts) < 3:
        return None
    if abs(float(dist_mm)) < 1e-6:
        return list(pts)
    poly = ensure_ccw(pts)
    n = len(poly)
    out = []
    d = float(dist_mm)
    for i in range(n):
        p_prev = poly[(i - 1) % n]
        p = poly[i]
        p_next = poly[(i + 1) % n]
        e1x = float(p[0]) - float(p_prev[0])
        e1y = float(p[1]) - float(p_prev[1])
        e2x = float(p_next[0]) - float(p[0])
        e2y = float(p_next[1]) - float(p[1])
        l1 = math.hypot(e1x, e1y)
        l2 = math.hypot(e2x, e2y)
        if l1 < _EPS or l2 < _EPS:
            out.append(p)
            continue
        n1x, n1y = -e1y / l1, e1x / l1
        n2x, n2y = -e2y / l2, e2x / l2
        bis_x = n1x + n2x
        bis_y = n1y + n2y
        bl = math.hypot(bis_x, bis_y)
        if bl < _EPS:
            out.append((float(p[0]) + n2x * d, float(p[1]) + n2y * d))
            continue
        bis_x /= bl
        bis_y /= bl
        c = max(_dot(n2x, n2y, bis_x, bis_y), 0.15)
        m = d / c
        if m > abs(d) * 4.0:
            m = abs(d) * 4.0
        out.append((float(p[0]) + bis_x * m, float(p[1]) + bis_y * m))
    cleaned = _clean_ring(out)
    if len(cleaned) < 3:
        return None
    a0 = shoelace_area_m2(poly)
    a1 = shoelace_area_m2(cleaned)
    # Colapso o offset inválido (área no se reduce).
    if a1 < 1e-6 or a1 >= a0:
        return None
    return ensure_ccw(cleaned)


def rect_from_two_points_mm(p1, p2):
    """
    Polígono rectangular en el plano a partir de 2 esquinas opuestas (mm).
    Ejes alineados al sistema del plano (X/Y del Sketch).
    """
    if p1 is None or p2 is None:
        return None
    x0 = min(float(p1[0]), float(p2[0]))
    x1 = max(float(p1[0]), float(p2[0]))
    y0 = min(float(p1[1]), float(p2[1]))
    y1 = max(float(p1[1]), float(p2[1]))
    if (x1 - x0) < 1.0 or (y1 - y0) < 1.0:
        return None
    return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]


def luz_menor_mm_from_polygon(pts):
    """
    Luz menor del paño = ``min(width, height)`` del AABB en mm.

    Misma base geométrica que ``span_direction_from_polygon_mm``.
    Returns:
        float mm, o None si degenerado / inválido.
    """
    if not pts or len(pts) < 3:
        return None
    xs = []
    ys = []
    for p in pts:
        try:
            xs.append(float(p[0]))
            ys.append(float(p[1]))
        except Exception:
            continue
    if len(xs) < 3:
        return None
    width = max(xs) - min(xs)
    height = max(ys) - min(ys)
    if width < _EPS or height < _EPS:
        return None
    return min(width, height)


def span_direction_from_polygon_mm(pts, equal_tol_mm=1.0):
    """
    Dirección Major/Principal en el plano (mm) según la luz menor del paño.

    Luz menor = ``min(width, height)`` del AABB del polígono en coords del
    plano Sketch (mm). Major = vector unitario a lo largo del lado corto:

    - ``width < height`` → horizontal ``(1, 0)``
    - ``height < width`` → vertical ``(0, 1)``
    - casi cuadrado (``|width - height| ≤ equal_tol_mm``) → horizontal
      estable ``(1, 0)`` (eje local del paño; no devolver ``None``)
    - degenerado (``width`` o ``height`` < eps) → ``None``

    Returns:
        (dx, dy) unitario en coords del plano Sketch, o None.
    """
    if not pts or len(pts) < 3:
        return None
    tol = float(equal_tol_mm)
    if tol < 0.0:
        tol = 0.0
    xs = []
    ys = []
    for p in pts:
        try:
            xs.append(float(p[0]))
            ys.append(float(p[1]))
        except Exception:
            continue
    if len(xs) < 3:
        return None
    width = max(xs) - min(xs)
    height = max(ys) - min(ys)
    if width < _EPS or height < _EPS:
        return None
    # Casi cuadrado: eje local estable (horizontal), sin fallback a la losa
    if abs(width - height) <= tol:
        return (1.0, 0.0)
    if width < height:
        return (1.0, 0.0)
    return (0.0, 1.0)


def point_in_polygon_mm(pt, poly, tol=_EPS_MM):
    """True si ``pt`` está dentro o en el borde de ``poly`` (mm)."""
    if not pt or not poly or len(poly) < 3:
        return False
    if _point_on_poly_boundary(pt, poly, tol):
        return True
    x, y = float(pt[0]), float(pt[1])
    inside = False
    n = len(poly)
    j = n - 1
    for i in range(n):
        xi, yi = float(poly[i][0]), float(poly[i][1])
        xj, yj = float(poly[j][0]), float(poly[j][1])
        if ((yi > y) != (yj > y)) and (
            x < (xj - xi) * (y - yi) / ((yj - yi) if abs(yj - yi) > _EPS else _EPS) + xi
        ):
            inside = not inside
        j = i
    return inside


def _seg_seg_distance_mm(a, b, c, d):
    """Distancia mínima entre segmentos ab y cd (mm)."""
    ax, ay = float(a[0]), float(a[1])
    bx, by = float(b[0]), float(b[1])
    cx, cy = float(c[0]), float(c[1])
    dx, dy = float(d[0]), float(d[1])
    hit = _line_seg_intersection(a, b, c, d)
    if hit is not None:
        return 0.0

    def _pt_seg(px, py, qx, qy, rx, ry):
        vx, vy = rx - qx, ry - qy
        L2 = vx * vx + vy * vy
        if L2 < _EPS:
            return math.hypot(px - qx, py - qy)
        t = max(0.0, min(1.0, ((px - qx) * vx + (py - qy) * vy) / L2))
        return math.hypot(px - (qx + t * vx), py - (qy + t * vy))

    return min(
        _pt_seg(ax, ay, cx, cy, dx, dy),
        _pt_seg(bx, by, cx, cy, dx, dy),
        _pt_seg(cx, cy, ax, ay, bx, by),
        _pt_seg(dx, dy, ax, ay, bx, by),
    )


def _poly_bbox(pts):
    xs = [float(p[0]) for p in pts]
    ys = [float(p[1]) for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


def polygons_overlap_or_touch_mm(a, b, tol=_EPS_MM):
    """True si dos anillos se solapan o comparten borde (tol en mm)."""
    if not a or not b or len(a) < 3 or len(b) < 3:
        return False
    t = float(tol)
    ax0, ay0, ax1, ay1 = _poly_bbox(a)
    bx0, by0, bx1, by1 = _poly_bbox(b)
    if ax1 < bx0 - t or bx1 < ax0 - t or ay1 < by0 - t or by1 < ay0 - t:
        return False
    for p in a:
        if point_in_polygon_mm(p, b, t):
            return True
    for p in b:
        if point_in_polygon_mm(p, a, t):
            return True
    na, nb = len(a), len(b)
    for i in range(na):
        a0, a1 = a[i], a[(i + 1) % na]
        for j in range(nb):
            b0, b1 = b[j], b[(j + 1) % nb]
            if _line_seg_intersection(a0, a1, b0, b1) is not None:
                return True
            if _seg_seg_distance_mm(a0, a1, b0, b1) <= t:
                return True
    return False


def polygons_form_single_component_mm(polygons, tol=_EPS_MM):
    """True si todos los polígonos forman un único componente por solape/toque."""
    polys = [p for p in (polygons or []) if p and len(p) >= 3]
    n = len(polys)
    if n <= 1:
        return n == 1
    parent = list(range(n))

    def _find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def _unite(i, j):
        ri, rj = _find(i), _find(j)
        if ri != rj:
            parent[rj] = ri

    for i in range(n):
        for j in range(i + 1, n):
            if polygons_overlap_or_touch_mm(polys[i], polys[j], tol):
                _unite(i, j)
    roots = set(_find(i) for i in range(n))
    return len(roots) == 1


def _is_orthogonal_ring(pts, tol=_EPS_MM):
    if not pts or len(pts) < 3:
        return False
    t = float(tol)
    n = len(pts)
    for i in range(n):
        x0, y0 = float(pts[i][0]), float(pts[i][1])
        x1, y1 = float(pts[(i + 1) % n][0]), float(pts[(i + 1) % n][1])
        if abs(x1 - x0) > t and abs(y1 - y0) > t:
            return False
    return True


def _unique_sorted_coords(vals, tol=_EPS_MM):
    if not vals:
        return []
    vals = sorted(float(v) for v in vals)
    out = [vals[0]]
    t = float(tol)
    for v in vals[1:]:
        if abs(v - out[-1]) > t:
            out.append(v)
    return out


def _trace_ortho_outer_ring(xs, ys, occ):
    """
    Contorno exterior CCW de celdas ocupadas en malla ortogonal.
    xs/ys: coordenadas de líneas de malla; occ[j][i] = celda (xs[i]..xs[i+1], ys[j]..ys[j+1]).
    """
    ny = len(occ)
    nx = len(occ[0]) if ny else 0
    if nx < 1 or ny < 1:
        return None

    edge_count = {}

    def _add_edge(i0, j0, i1, j1):
        key = (i0, j0, i1, j1)
        rev = (i1, j1, i0, j0)
        if rev in edge_count and edge_count[rev] > 0:
            edge_count[rev] -= 1
            if edge_count[rev] == 0:
                del edge_count[rev]
        else:
            edge_count[key] = edge_count.get(key, 0) + 1

    for j in range(ny):
        for i in range(nx):
            if not occ[j][i]:
                continue
            if j == 0 or not occ[j - 1][i]:
                _add_edge(i, j, i + 1, j)
            if i == nx - 1 or not occ[j][i + 1]:
                _add_edge(i + 1, j, i + 1, j + 1)
            if j == ny - 1 or not occ[j + 1][i]:
                _add_edge(i + 1, j + 1, i, j + 1)
            if i == 0 or not occ[j][i - 1]:
                _add_edge(i, j + 1, i, j)

    out_map = {}
    for (i0, j0, i1, j1), cnt in edge_count.items():
        if cnt <= 0:
            continue
        out_map.setdefault((i0, j0), []).append((i1, j1))
    if not out_map:
        return None

    # Arranque: arista inferior más baja (luego más a la izquierda) hacia la derecha
    start = None
    first_to = None
    best_key = None
    for (i0, j0), dests in out_map.items():
        for (i1, j1) in dests:
            if j0 != j1 or i1 != i0 + 1:
                continue
            key = (j0, i0)
            if best_key is None or key < best_key:
                best_key = key
                start = (i0, j0)
                first_to = (i1, j1)
    if start is None:
        # Fallback: cualquier arista
        start = next(iter(out_map.keys()))
        first_to = out_map[start][0]

    def _pop_best(node, prev):
        options = out_map.get(node) or []
        if not options:
            return None
        px, py = prev
        cx, cy = node
        in_dx, in_dy = cx - px, cy - py

        def _turn_key(dest):
            dx = dest[0] - cx
            dy = dest[1] - cy
            cross = in_dx * dy - in_dy * dx
            dot = in_dx * dx + in_dy * dy
            if cross > 0:
                rank = 0
            elif abs(cross) < _EPS and dot > 0:
                rank = 1
            elif cross < 0:
                rank = 2
            else:
                rank = 3
            return (rank, -dot)

        options_sorted = sorted(options, key=_turn_key)
        chosen = options_sorted[0]
        options.remove(chosen)
        if not options:
            del out_map[node]
        else:
            out_map[node] = options
        return chosen

    # Consumir la arista inicial
    init_opts = out_map.get(start) or []
    if first_to in init_opts:
        init_opts.remove(first_to)
        if not init_opts:
            del out_map[start]
        else:
            out_map[start] = init_opts
    else:
        return None

    ring_idx = [start, first_to]
    cur = start
    nxt = first_to
    guard = 0
    max_g = (nx + 1) * (ny + 1) * 4 + 8
    while nxt != start and guard < max_g:
        nxt2 = _pop_best(nxt, cur)
        if nxt2 is None:
            return None
        cur = nxt
        nxt = nxt2
        ring_idx.append(nxt)
        guard += 1
    if ring_idx[-1] != start or len(ring_idx) < 4:
        return None
    ring_idx = ring_idx[:-1]
    pts = [(float(xs[i]), float(ys[j])) for (i, j) in ring_idx]
    return ensure_ccw(_clean_ring(pts))


def union_polygons_mm(polygons, tol=_EPS_MM):
    """
    Unión booleana de polígonos en mm.

    Requiere que se solapen o toquen (un solo componente). Preferencia:
    malla ortogonal (rectángulos / anillos de aristas eje-alineadas).

    Returns:
        anillo CCW (lista de pts) o None si falla / disjuntos / degenerado.
    """
    polys = []
    for p in polygons or []:
        ring = ensure_ccw(_clean_ring(p))
        if ring and len(ring) >= 3 and shoelace_area_m2(ring) > 1e-8:
            polys.append(ring)
    if not polys:
        return None
    if len(polys) == 1:
        return list(polys[0])
    if not polygons_form_single_component_mm(polys, tol):
        return None

    if all(_is_orthogonal_ring(p, tol) for p in polys):
        xs = []
        ys = []
        for p in polys:
            for q in p:
                xs.append(float(q[0]))
                ys.append(float(q[1]))
        xs = _unique_sorted_coords(xs, tol)
        ys = _unique_sorted_coords(ys, tol)
        nx = len(xs) - 1
        ny = len(ys) - 1
        if nx < 1 or ny < 1:
            return None
        occ = [[False] * nx for _ in range(ny)]
        any_occ = False
        for j in range(ny):
            for i in range(nx):
                cx = 0.5 * (xs[i] + xs[i + 1])
                cy = 0.5 * (ys[j] + ys[j + 1])
                for poly in polys:
                    if point_in_polygon_mm((cx, cy), poly, tol):
                        occ[j][i] = True
                        any_occ = True
                        break
        if not any_occ:
            return None
        # Conectividad 4 de celdas
        visited = [[False] * nx for _ in range(ny)]
        start = None
        for j in range(ny):
            for i in range(nx):
                if occ[j][i]:
                    start = (i, j)
                    break
            if start is not None:
                break
        stack = [start]
        visited[start[1]][start[0]] = True
        count = 0
        total = sum(1 for j in range(ny) for i in range(nx) if occ[j][i])
        while stack:
            ci, cj = stack.pop()
            count += 1
            for di, dj in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                ni, nj = ci + di, cj + dj
                if ni < 0 or nj < 0 or ni >= nx or nj >= ny:
                    continue
                if not occ[nj][ni] or visited[nj][ni]:
                    continue
                visited[nj][ni] = True
                stack.append((ni, nj))
        if count != total:
            return None
        ring = _trace_ortho_outer_ring(xs, ys, occ)
        if not ring or len(ring) < 3:
            return None
        ring = ensure_ccw(_remove_collinear_mm(_clean_ring(ring), tol))
        if len(ring) < 3 or shoelace_area_m2(ring) <= 1e-8:
            return None
        return ring

    # Fallback no ortogonal: unión sucesiva vía inserción de cortes (solo solapes
    # parciales simples). Si no se puede, None.
    return _union_polygons_general_mm(polys, tol)


def _union_polygons_general_mm(polys, tol=_EPS_MM):
    """
    Unión general limitada: si un polígono contiene a otro, conserva el mayor;
    si no, intenta unión ortogonal del AABB de cada uno solo cuando los AABB
    bastan (no). Mejor: fusionar por Combined no disponible aquí.

    Estrategia: ir acumulando con unión ortogonal de los AABB *solo si* cada
    anillo es ya ortogonal (ya filtrado). Para no-ortogonal, devolver el
    envolvente convexo solo no — preferimos None.
    """
    # Acumular: A ∪ B = A si B ⊂ A, B si A ⊂ B; si se cruzan sin ortogonalidad → None
    acc = list(polys[0])
    for other in polys[1:]:
        a_in_b = all(point_in_polygon_mm(p, other, tol) for p in acc)
        b_in_a = all(point_in_polygon_mm(p, acc, tol) for p in other)
        if b_in_a and not a_in_b:
            continue
        if a_in_b and not b_in_a:
            acc = list(other)
            continue
        if a_in_b and b_in_a:
            # iguales / casi
            if shoelace_area_m2(other) > shoelace_area_m2(acc):
                acc = list(other)
            continue
        # Cruce verdadero sin soporte ortogonal
        return None
    return ensure_ccw(_clean_ring(acc))


def clip_polygon_halfplane_mm(pts, nx, ny, min_dot):
    """
    Sutherland–Hodgman: conserva el semiplano ``p·n̂ >= min_dot`` (n se normaliza).
    """
    if not pts or len(pts) < 3:
        return None
    ln = math.hypot(float(nx), float(ny))
    if ln < _EPS:
        return None
    ux = float(nx) / ln
    uy = float(ny) / ln
    limit = float(min_dot)

    def _inside(p):
        return (float(p[0]) * ux + float(p[1]) * uy) >= (limit - 1e-6)

    def _hit(a, b):
        ax, ay = float(a[0]), float(a[1])
        bx, by = float(b[0]), float(b[1])
        da = ax * ux + ay * uy
        db = bx * ux + by * uy
        denom = db - da
        if abs(denom) < _EPS:
            return (ax, ay)
        t = (limit - da) / denom
        if t < 0.0:
            t = 0.0
        elif t > 1.0:
            t = 1.0
        return (ax + t * (bx - ax), ay + t * (by - ay))

    out = []
    n = len(pts)
    for i in range(n):
        s = pts[(i - 1) % n]
        e = pts[i]
        ein = _inside(e)
        sin = _inside(s)
        if ein:
            if not sin:
                out.append(_hit(s, e))
            out.append((float(e[0]), float(e[1])))
        elif sin:
            out.append(_hit(s, e))
    cleaned = ensure_ccw(_clean_ring(out))
    if len(cleaned) < 3 or shoelace_area_m2(cleaned) < 1e-8:
        return None
    return cleaned


def cutback_polygon_end_mm(pts, ux, uy, cut_mm, which_end):
    """
    Recorta una franja ``cut_mm`` en un extremo del eje ``(ux,uy)``.

    ``which_end``: 0 = proyección baja (start), 1 = proyección alta (end).
    Devuelve ``None`` si el recorte es inválido o colapsa el polígono.
    """
    if not pts or len(pts) < 3:
        return None
    cut = float(cut_mm)
    if cut <= 0.5:
        return ensure_ccw([(float(p[0]), float(p[1])) for p in pts])
    ln = math.hypot(float(ux), float(uy))
    if ln < _EPS:
        return None
    bx = float(ux) / ln
    by = float(uy) / ln
    projs = [float(p[0]) * bx + float(p[1]) * by for p in pts]
    pmin = min(projs)
    pmax = max(projs)
    span = pmax - pmin
    if span < 1.0 or cut >= span * 0.45:
        return None
    if int(which_end) == 0:
        return clip_polygon_halfplane_mm(pts, bx, by, pmin + cut)
    return clip_polygon_halfplane_mm(pts, -bx, -by, -(pmax - cut))


def _snap_array_length_to_spacing_multiple(length_mm, spacing_mm):
    """
    ``L' = floor(L / s) * s`` (nunca crece).

    Con Layout Rule = Maximum Spacing, el paso real es exacto solo si el largo
    de layout es múltiplo de ``s``. Devuelve ``None`` si ``L' < s``.
    """
    try:
        L = float(length_mm)
        s = float(spacing_mm)
    except Exception:
        return None
    if L < 1.0 or s < 1.0:
        return None
    # +1e-9 evita que L casi múltiplo (p.ej. 2999.999999) baje un intervalo
    n = int(math.floor((L / s) + 1e-9))
    if n < 1:
        return None
    return float(n) * s


def snap_polygon_dist_span_to_spacing_mm(pts, dist_ux, dist_uy, spacing_mm):
    """
    Encoge el polígono en el eje de distribución hasta ``L' = floor(L/s)*s``.

    El exceso se quita del extremo de proyección alta, de modo que el borde bajo
    —origen de layout / inset e del Set B— se conserva. Clip directo (no usa el
    tope 45% de ``cutback_polygon_end_mm``: el exceso puede acercarse a ``s``).
    No crece fuera del host: solo reduce. ``None`` si ``L' < s`` o el clip falla.
    """
    if not pts or len(pts) < 3:
        return None
    ln = math.hypot(float(dist_ux), float(dist_uy))
    if ln < _EPS:
        return None
    dx = float(dist_ux) / ln
    dy = float(dist_uy) / ln
    s = float(spacing_mm)
    if s < 1.0:
        return None
    projs = [float(p[0]) * dx + float(p[1]) * dy for p in pts]
    pmin = min(projs)
    pmax = max(projs)
    L = pmax - pmin
    L_prime = _snap_array_length_to_spacing_multiple(L, s)
    if L_prime is None:
        return None
    excess = L - L_prime
    if excess <= 0.5:
        return ensure_ccw([(float(p[0]), float(p[1])) for p in pts])
    # Conservar pmin; nuevo pmax = pmin + L' (solo eje distribución)
    return clip_polygon_halfplane_mm(pts, -dx, -dy, -(pmin + L_prime))


def ahorro_fierro_polygons_mm(pts, bar_ux, bar_uy, spacing_e_mm, cutback_pct=10.0):
    """
    Polígonos Set A / Set B para «ahorro de fierro» (series intercaladas).

    Pedido Ø @ e → cada set con espaciamiento 2e (lo aplica el llamador).
    - Set A: recorte ``cutback_pct`` % de la luz en el extremo start del eje barra.
    - Set B: recorte en el extremo opuesto + inset ``e`` desde el lado bajo de la
      dirección de distribución (perpendicular) para intercalado visual a e.
    - Tras cutbacks/inset: luz de distribución de cada set se redondea a
      ``floor(L / 2e) * 2e`` (encoge borde alto) para que Maximum Spacing dé
      paso exacto ``2e`` sin disolver el AreaReinforcement.

    Returns:
        ``[(u"A", pts_A), (u"B", pts_B)]`` o ``None`` si no se puede aplicar.
    """
    if not pts or len(pts) < 3:
        return None
    ln = math.hypot(float(bar_ux), float(bar_uy))
    if ln < _EPS:
        return None
    ux = float(bar_ux) / ln
    uy = float(bar_uy) / ln
    # Distribución de barras = perpendicular al eje de la barra
    dx, dy = -uy, ux
    e = float(spacing_e_mm)
    if e < 1.0:
        return None
    # Espaciado real de cada AR intercalado (Maximum Spacing = 2e)
    s_set = 2.0 * e
    pct = float(cutback_pct)
    if pct <= 0.0:
        pct = 10.0
    projs_bar = [float(p[0]) * ux + float(p[1]) * uy for p in pts]
    span_bar = max(projs_bar) - min(projs_bar)
    if span_bar < 1.0:
        return None
    cut = span_bar * (pct / 100.0)
    if cut < 1.0:
        cut = 1.0
    pts_a = cutback_polygon_end_mm(pts, ux, uy, cut, 0)
    pts_b = cutback_polygon_end_mm(pts, ux, uy, cut, 1)
    if pts_a is None or pts_b is None:
        return None
    # Intercalado Set B: desplazar el origen de layout ~e en dirección de distribución
    pts_b = cutback_polygon_end_mm(pts_b, dx, dy, e, 0)
    if pts_b is None:
        return None
    # Snap dist. tras inset B: L' = floor(L/s)*s con s=2e (exact spacing sin dissolve)
    pts_a = snap_polygon_dist_span_to_spacing_mm(pts_a, dx, dy, s_set)
    pts_b = snap_polygon_dist_span_to_spacing_mm(pts_b, dx, dy, s_set)
    if pts_a is None or pts_b is None:
        return None
    if shoelace_area_m2(pts_a) < 1e-8 or shoelace_area_m2(pts_b) < 1e-8:
        return None
    return [(u"A", pts_a), (u"B", pts_b)]
