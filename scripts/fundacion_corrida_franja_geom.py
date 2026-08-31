# -*- coding: utf-8 -*-
"""
Geometría de franjas (polígonos) → ejes long./trans. para fundación corrida.

Contornos de host en coordenadas de la **vista activa** (Right/Up), no AABB mundo.

Revit 2024–2026 · IronPython.
"""

from __future__ import print_function

import math

from Autodesk.Revit.DB import FilteredElementCollector, Line, WallFoundation, XYZ

_FT_TO_MM = 304.8
_MM_TO_FT = 1.0 / 304.8


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def collect_wall_foundations_in_view(doc, view):
    """Lista de ``WallFoundation`` visibles en ``view``."""
    out = []
    if doc is None or view is None:
        return out
    try:
        col = FilteredElementCollector(doc, view.Id).OfClass(WallFoundation)
        for el in col:
            if el is not None and isinstance(el, WallFoundation):
                out.append(el)
    except Exception:
        try:
            col = FilteredElementCollector(doc).OfClass(WallFoundation)
            for el in col:
                if el is None or not isinstance(el, WallFoundation):
                    continue
                try:
                    bb = el.get_BoundingBox(view)
                except Exception:
                    bb = None
                if bb is not None:
                    out.append(el)
        except Exception:
            pass
    return out


def view_frame_from_view(view):
    """
    Marco 2D de la vista de planta: origen + Right/Up (mundo).

    Coordenadas de canvas: u = RightDirection, v = UpDirection (mm).
    """
    if view is None:
        return None
    try:
        o = view.Origin
        r = view.RightDirection
        u = view.UpDirection
        ox, oy, oz = float(o.X), float(o.Y), float(o.Z)
        rx, ry, rz = float(r.X), float(r.Y), float(r.Z)
        ux, uy, uz = float(u.X), float(u.Y), float(u.Z)
        rlen = math.sqrt(rx * rx + ry * ry + rz * rz)
        ulen = math.sqrt(ux * ux + uy * uy + uz * uz)
        if rlen < 1e-12 or ulen < 1e-12:
            return None
        rx, ry, rz = rx / rlen, ry / rlen, rz / rlen
        ux, uy, uz = ux / ulen, uy / ulen, uz / ulen
        return {
            u"ox": ox,
            u"oy": oy,
            u"oz": oz,
            u"rx": rx,
            u"ry": ry,
            u"rz": rz,
            u"ux": ux,
            u"uy": uy,
            u"uz": uz,
        }
    except Exception:
        return None


def world_xyz_to_view_mm(x, y, z, frame):
    if frame is None:
        return (float(x) * _FT_TO_MM, float(y) * _FT_TO_MM)
    dx = float(x) - float(frame[u"ox"])
    dy = float(y) - float(frame[u"oy"])
    dz = float(z) - float(frame[u"oz"])
    uu = (
        dx * float(frame[u"rx"])
        + dy * float(frame[u"ry"])
        + dz * float(frame[u"rz"])
    ) * _FT_TO_MM
    vv = (
        dx * float(frame[u"ux"])
        + dy * float(frame[u"uy"])
        + dz * float(frame[u"uz"])
    ) * _FT_TO_MM
    return (uu, vv)


def view_mm_to_world_xy_ft(u_mm, v_mm, frame, z_ft=None):
    """Punto mundo (X,Y,Z) desde UV de vista (mm)."""
    if frame is None:
        z = 0.0 if z_ft is None else float(z_ft)
        return (
            float(u_mm) * _MM_TO_FT,
            float(v_mm) * _MM_TO_FT,
            z,
        )
    u = float(u_mm) * _MM_TO_FT
    v = float(v_mm) * _MM_TO_FT
    x = float(frame[u"ox"]) + u * float(frame[u"rx"]) + v * float(frame[u"ux"])
    y = float(frame[u"oy"]) + u * float(frame[u"ry"]) + v * float(frame[u"uy"])
    z = float(frame[u"oz"]) + u * float(frame[u"rz"]) + v * float(frame[u"uz"])
    if z_ft is not None:
        z = float(z_ft)
    return (x, y, z)


def view_dir_mm_to_world_xy(du, dv, frame):
    """Dirección 2D en vista (du,dv) → (dx,dy) mundo XY (sin Z)."""
    if frame is None:
        return float(du), float(dv)
    dx = float(du) * float(frame[u"rx"]) + float(dv) * float(frame[u"ux"])
    dy = float(du) * float(frame[u"ry"]) + float(dv) * float(frame[u"uy"])
    return dx, dy


def _sample_curve_xyz(curve, n_seg=16):
    pts = []
    if curve is None:
        return pts
    try:
        n = max(2, int(n_seg))
        for i in range(n + 1):
            t = float(i) / float(n)
            try:
                p = curve.Evaluate(t, True)
            except Exception:
                if i == 0:
                    p = curve.GetEndPoint(0)
                elif i == n:
                    p = curve.GetEndPoint(1)
                else:
                    continue
            pts.append((float(p.X), float(p.Y), float(p.Z)))
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            pts = [
                (float(p0.X), float(p0.Y), float(p0.Z)),
                (float(p1.X), float(p1.Y), float(p1.Z)),
            ]
        except Exception:
            pass
    return pts


def _loop_to_xyz_points(curve_loop):
    pts = []
    try:
        from geometria_fundacion_cara_inferior import _iter_curvas_en_curveloop
    except Exception:
        return pts
    for c in _iter_curvas_en_curveloop(curve_loop):
        seg = _sample_curve_xyz(c, n_seg=20)
        if not seg:
            continue
        if pts:
            last = pts[-1]
            first = seg[0]
            if (
                abs(last[0] - first[0]) < 1e-4
                and abs(last[1] - first[1]) < 1e-4
                and abs(last[2] - first[2]) < 1e-4
            ):
                seg = seg[1:]
        pts.extend(seg)
    if len(pts) >= 3:
        first, last = pts[0], pts[-1]
        if (
            abs(first[0] - last[0]) < 1e-4
            and abs(first[1] - last[1]) < 1e-4
            and abs(first[2] - last[2]) < 1e-4
        ):
            pts = pts[:-1]
    return pts if len(pts) >= 3 else []


def extract_wall_foundation_outline_xyz(wf):
    """
    Contorno(s) 3D reales de la zapata (cara inferior / bbox).

    Returns: list[list[(x,y,z)]] en pies internos.
    """
    out = []
    if wf is None:
        return out
    try:
        from geometria_fundacion_cara_inferior import (
            elegir_loop_mayor_perimetro,
            extraer_curvas_perimetrales_cara_inferior,
        )

        r = extraer_curvas_perimetrales_cara_inferior(wf)
    except Exception:
        r = None
    if r is not None:
        loops, _z = r
        ordered = []
        try:
            best = elegir_loop_mayor_perimetro(loops) if loops else None
        except Exception:
            best = None
        if best is not None:
            ordered.append(best)
            for cl in loops or []:
                if cl is best:
                    continue
                ordered.append(cl)
        else:
            ordered = list(loops or [])
        for cl in ordered:
            pts = _loop_to_xyz_points(cl)
            if pts:
                out.append(pts)
    if out:
        return out
    try:
        bb = wf.get_BoundingBox(None)
        if bb is None:
            return out
        z = 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
        x0, y0 = float(bb.Min.X), float(bb.Min.Y)
        x1, y1 = float(bb.Max.X), float(bb.Max.Y)
        out.append(
            [
                (x0, y0, z),
                (x1, y0, z),
                (x1, y1, z),
                (x0, y1, z),
            ]
        )
    except Exception:
        pass
    return out


def _view_crop_uv_mm(view, frame):
    """Rectángulo de recorte en UV mm: (u_min, u_max, v_min, v_max) o None."""
    if view is None or frame is None:
        return None
    try:
        if not bool(view.CropBoxActive):
            return None
    except Exception:
        return None
    try:
        cb = view.CropBox
        if cb is None:
            return None
        tr = cb.Transform
        corners = []
        for xv in (float(cb.Min.X), float(cb.Max.X)):
            for yv in (float(cb.Min.Y), float(cb.Max.Y)):
                p = tr.OfPoint(XYZ(xv, yv, float(cb.Min.Z)))
                corners.append(
                    world_xyz_to_view_mm(
                        float(p.X), float(p.Y), float(p.Z), frame
                    )
                )
        if not corners:
            return None
        us = [c[0] for c in corners]
        vs = [c[1] for c in corners]
        return (min(us), max(us), min(vs), max(vs))
    except Exception:
        return None


def build_host_preview_in_view(wf, view):
    """
    Preview fiel a la vista: contorno real proyectado a Right/Up.

    Returns dict:
      poly, polys, view_frame, label, length_mm, width_mm, crop_uv
    o None.
    """
    if wf is None or not isinstance(wf, WallFoundation):
        return None
    frame = view_frame_from_view(view)
    if frame is None:
        return None
    loops_xyz = extract_wall_foundation_outline_xyz(wf)
    if not loops_xyz:
        try:
            bb = wf.get_BoundingBox(view)
        except Exception:
            bb = None
        if bb is not None:
            z = 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
            loops_xyz = [
                [
                    (float(bb.Min.X), float(bb.Min.Y), z),
                    (float(bb.Max.X), float(bb.Min.Y), z),
                    (float(bb.Max.X), float(bb.Max.Y), z),
                    (float(bb.Min.X), float(bb.Max.Y), z),
                ]
            ]
    if not loops_xyz:
        return None
    crop = _view_crop_uv_mm(view, frame)
    polys = []
    for loop in loops_xyz:
        poly = [
            world_xyz_to_view_mm(p[0], p[1], p[2], frame) for p in loop
        ]
        if poly and len(poly) >= 3:
            polys.append(poly)
    if not polys:
        return None
    xs = []
    ys = []
    for poly in polys:
        for x, y in poly:
            xs.append(float(x))
            ys.append(float(y))
    length_mm = max(xs) - min(xs) if xs else 0.0
    width_mm = max(ys) - min(ys) if ys else 0.0
    if width_mm > length_mm:
        length_mm, width_mm = width_mm, length_mm
    label = u"Wall Foundation"
    try:
        from barras_bordes_losa_gancho_empotramiento import element_id_to_int

        eid = element_id_to_int(wf.Id)
        if eid is not None:
            label = u"Wall Foundation Id {0}".format(eid)
    except Exception:
        pass
    try:
        nm = _as_unicode(wf.Name)
        if nm:
            label = u"{0} · {1}".format(nm, label)
    except Exception:
        pass
    return {
        u"poly": polys[0],
        u"polys": polys,
        u"view_frame": frame,
        u"label": label,
        u"length_mm": float(length_mm),
        u"width_mm": float(width_mm),
        u"crop_uv": crop,
    }


def polygon_centroid_mm(poly):
    if not poly:
        return None
    n = len(poly)
    if n < 1:
        return None
    sx = sy = 0.0
    for x, y in poly:
        sx += float(x)
        sy += float(y)
    return (sx / float(n), sy / float(n))


def point_in_polygon_mm(px, py, poly):
    """Ray casting; poly cerrado o abierto."""
    if not poly or len(poly) < 3:
        return False
    n = len(poly)
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = float(poly[i][0]), float(poly[i][1])
        xj, yj = float(poly[j][0]), float(poly[j][1])
        if ((yi > py) != (yj > py)) and (
            px < (xj - xi) * (py - yi) / ((yj - yi) + 1e-30) + xi
        ):
            inside = not inside
        j = i
    return inside


def strip_axes_from_polygon_mm(poly):
    """
    Franja: lado más largo → eje longitudinal; proyección ⟂ → ancho.
    """
    if not poly or len(poly) < 3:
        return None
    pts = [(float(p[0]), float(p[1])) for p in poly]
    n = len(pts)
    best = None
    for i in range(n):
        x0, y0 = pts[i]
        x1, y1 = pts[(i + 1) % n]
        dx, dy = x1 - x0, y1 - y0
        L = math.hypot(dx, dy)
        if L < 1.0:
            continue
        if best is None or L > best[0]:
            best = (L, dx / L, dy / L, x0, y0, x1, y1)
    if best is None:
        return None
    L_long, ux, uy = best[0], best[1], best[2]
    vx, vy = -uy, ux
    us = []
    vs = []
    ox, oy = pts[0]
    for x, y in pts:
        us.append((x - ox) * ux + (y - oy) * uy)
        vs.append((x - ox) * vx + (y - oy) * vy)
    u_min, u_max = min(us), max(us)
    v_min, v_max = min(vs), max(vs)
    length_mm = max(1.0, u_max - u_min)
    width_mm = max(1.0, v_max - v_min)
    if width_mm > length_mm + 1.0:
        ux, uy, vx, vy = vx, vy, -ux, -uy
        u_min, u_max, v_min, v_max = v_min, v_max, u_min, u_max
        length_mm, width_mm = width_mm, length_mm
    u_mid = 0.5 * (u_min + u_max)
    v_mid = 0.5 * (v_min + v_max)

    def to_local(u, v):
        return (ox + u * ux + v * vx, oy + u * uy + v * vy)

    p0 = to_local(u_min, v_mid)
    p1 = to_local(u_max, v_mid)
    center = to_local(u_mid, v_mid)
    return {
        u"poly": list(pts),
        u"length_mm": float(length_mm),
        u"width_mm": float(width_mm),
        u"p0_mm": p0,
        u"p1_mm": p1,
        u"center_mm": center,
        u"u_hat": (ux, uy),
        u"v_hat": (vx, vy),
        u"u_min": u_min,
        u"u_max": u_max,
        u"v_min": v_min,
        u"v_max": v_max,
        u"origin_mm": (ox, oy),
    }


def assign_host_to_strip(hosts_with_preview, strip):
    if not hosts_with_preview or not strip:
        return None
    c = strip.get(u"center_mm")
    if c is None:
        c = polygon_centroid_mm(strip.get(u"poly") or [])
    if c is None:
        return None
    cx, cy = float(c[0]), float(c[1])
    for wf, prev in hosts_with_preview:
        if prev is None:
            continue
        for poly in prev.get(u"polys") or [prev.get(u"poly")]:
            if poly and point_in_polygon_mm(cx, cy, poly):
                return wf
    best = None
    best_d = None
    for wf, prev in hosts_with_preview:
        if prev is None:
            continue
        pc = polygon_centroid_mm(prev.get(u"poly") or [])
        if pc is None:
            continue
        d = math.hypot(cx - float(pc[0]), cy - float(pc[1]))
        if best_d is None or d < best_d:
            best_d = d
            best = wf
    if best is not None:
        return best
    try:
        return hosts_with_preview[0][0]
    except Exception:
        return None


def mm_to_ft(mm):
    return float(mm) * _MM_TO_FT


def ft_to_mm(ft):
    return float(ft) * _FT_TO_MM


_SNAP_KIND_VERTEX = u"vertex"
_SNAP_KIND_MID = u"mid"
_SNAP_KIND_EDGE = u"edge"


def _append_ring_snap_targets(targets, poly, closed=True):
    """Vértices, medios y aristas de un anillo (o polilínea si ``closed=False``)."""
    if not poly or len(poly) < 2:
        return
    pts = [(float(p[0]), float(p[1])) for p in poly]
    n = len(pts)
    for i in range(n):
        x0, y0 = pts[i]
        targets.append({u"kind": _SNAP_KIND_VERTEX, u"x": x0, u"y": y0})
    seg_count = n if closed else n - 1
    for i in range(seg_count):
        x0, y0 = pts[i]
        x1, y1 = pts[(i + 1) % n] if closed else pts[i + 1]
        dx, dy = x1 - x0, y1 - y0
        L = math.hypot(dx, dy)
        if L < 1.0:
            continue
        mx, my = 0.5 * (x0 + x1), 0.5 * (y0 + y1)
        targets.append({u"kind": _SNAP_KIND_MID, u"x": mx, u"y": my})
        targets.append(
            {
                u"kind": _SNAP_KIND_EDGE,
                u"x": x0,
                u"y": y0,
                u"x1": x1,
                u"y1": y1,
            }
        )


def build_foundation_snap_targets(previews):
    """Compatibilidad: solo contornos de host."""
    return build_canvas_snap_targets(previews=previews)


def build_canvas_snap_targets(previews, strips=None, draft_pts=None):
    """
    Objetivos OSNAP del canvas: zapatas, franjas cerradas, borrador y recorte.

    Permite colocar vértices sobre aristas (proyección), no solo en esquinas/medios.
    """
    targets = []
    for prev in previews or []:
        if not prev:
            continue
        polys = prev.get(u"polys")
        if not polys:
            p0 = prev.get(u"poly")
            polys = [p0] if p0 else []
        for poly in polys:
            _append_ring_snap_targets(targets, poly, closed=True)
        crop = prev.get(u"crop_uv")
        if crop is not None:
            try:
                u0, u1, v0, v1 = crop
                crop_poly = [
                    (float(u0), float(v0)),
                    (float(u1), float(v0)),
                    (float(u1), float(v1)),
                    (float(u0), float(v1)),
                ]
                _append_ring_snap_targets(targets, crop_poly, closed=True)
            except Exception:
                pass
    for strip in strips or []:
        poly = strip.get(u"poly") if strip else None
        if poly:
            _append_ring_snap_targets(targets, poly, closed=True)
    draft = list(draft_pts or [])
    if len(draft) == 1:
        try:
            x0, y0 = float(draft[0][0]), float(draft[0][1])
            targets.append({u"kind": _SNAP_KIND_VERTEX, u"x": x0, u"y": y0})
        except Exception:
            pass
    elif len(draft) >= 2:
        _append_ring_snap_targets(targets, draft, closed=False)
    return targets


def _dist2(ax, ay, bx, by):
    dx = float(ax) - float(bx)
    dy = float(ay) - float(by)
    return dx * dx + dy * dy


def _project_point_on_segment(px, py, x0, y0, x1, y1):
    dx = float(x1) - float(x0)
    dy = float(y1) - float(y0)
    L2 = dx * dx + dy * dy
    if L2 < 1e-12:
        return float(x0), float(y0), 0.0
    t = ((float(px) - float(x0)) * dx + (float(py) - float(y0)) * dy) / L2
    if t < 0.0:
        t = 0.0
    elif t > 1.0:
        t = 1.0
    return float(x0) + t * dx, float(y0) + t * dy, t


def snap_point_to_foundations_mm(px, py, targets, tol_mm):
    if not targets or tol_mm is None or float(tol_mm) <= 0:
        return float(px), float(py), None
    tol = float(tol_mm)
    tol2 = tol * tol
    best = None
    for t in targets:
        kind = t.get(u"kind")
        if kind == _SNAP_KIND_EDGE:
            qx, qy, _u = _project_point_on_segment(
                px, py, t[u"x"], t[u"y"], t[u"x1"], t[u"y1"]
            )
            d2 = _dist2(px, py, qx, qy)
            if d2 > tol2:
                continue
            pri = 2
            cand = (pri, d2, qx, qy, kind)
        else:
            qx, qy = float(t[u"x"]), float(t[u"y"])
            d2 = _dist2(px, py, qx, qy)
            if d2 > tol2:
                continue
            pri = 0 if kind == _SNAP_KIND_VERTEX else 1
            cand = (pri, d2, qx, qy, kind)
        if best is None:
            best = cand
            continue
        if cand[0] < best[0] or (cand[0] == best[0] and cand[1] < best[1]):
            best = cand
    if best is None:
        return float(px), float(py), None
    return best[2], best[3], best[4]


def build_lines_from_strip_ft(strip, z_ft, view_frame=None):
    """
    ``Line`` long y width en pies internos a cota ``z_ft``.

    ``strip`` en UV de vista (mm); ``view_frame`` convierte a mundo.
    """
    if strip is None:
        return None, None, 0.0
    p0m = strip.get(u"p0_mm")
    p1m = strip.get(u"p1_mm")
    vh = strip.get(u"v_hat")
    w_mm = float(strip.get(u"width_mm") or 0.0)
    if not p0m or not p1m or not vh or w_mm < 1.0:
        return None, None, 0.0
    z = float(z_ft)
    x0, y0, _z0 = view_mm_to_world_xy_ft(p0m[0], p0m[1], view_frame, z)
    x1, y1, _z1 = view_mm_to_world_xy_ft(p1m[0], p1m[1], view_frame, z)
    p0 = XYZ(x0, y0, z)
    p1 = XYZ(x1, y1, z)
    try:
        long_line = Line.CreateBound(p0, p1)
    except Exception:
        return None, None, 0.0
    mid = XYZ(
        0.5 * (float(p0.X) + float(p1.X)),
        0.5 * (float(p0.Y) + float(p1.Y)),
        z,
    )
    half = 0.5 * mm_to_ft(w_mm)
    vdx, vdy = view_dir_mm_to_world_xy(float(vh[0]), float(vh[1]), view_frame)
    vlen = math.hypot(vdx, vdy)
    if vlen < 1e-12:
        return long_line, None, 0.0
    vdx /= vlen
    vdy /= vlen
    wa = XYZ(float(mid.X) - vdx * half, float(mid.Y) - vdy * half, z)
    wb = XYZ(float(mid.X) + vdx * half, float(mid.Y) + vdy * half, z)
    try:
        width_line = Line.CreateBound(wa, wb)
    except Exception:
        return long_line, None, 0.0
    return long_line, width_line, float(width_line.Length)


def merge_strip_into_host_geo(base_geo, strip, view_frame=None):
    if base_geo is None or strip is None:
        return None
    geo = dict(base_geo)
    z_ref = None
    try:
        ll = geo.get(u"long_line") or geo.get("long_line")
        if ll is not None:
            z_ref = 0.5 * (
                float(ll.GetEndPoint(0).Z) + float(ll.GetEndPoint(1).Z)
            )
    except Exception:
        z_ref = None
    if z_ref is None:
        try:
            z0 = geo.get("z0")
            z1 = geo.get("z1")
            if z0 is not None and z1 is not None:
                z_ref = 0.5 * (float(z0) + float(z1))
        except Exception:
            z_ref = 0.0
    long_line, width_line, usable_w = build_lines_from_strip_ft(
        strip, z_ref or 0.0, view_frame=view_frame
    )
    if long_line is None or width_line is None:
        return None
    geo["long_line"] = long_line
    geo["width_line"] = width_line
    geo["usable_w_ft"] = float(usable_w)
    return geo
