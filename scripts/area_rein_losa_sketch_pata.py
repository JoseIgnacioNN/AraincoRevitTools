# -*- coding: utf-8 -*-
"""
Patas L geométricas para Area Rein. Losa Sketch.

- Largo = espesor losa − 50 mm (``largo_pata_mm_desde_espesor_host``).
- Inferior → +normal (arriba); Superior → −normal (abajo).
- Detección: aristas del paño (perímetro + cortes shaft/hueco) que
  **coinciden** con outline de losa o con shafts/huecos Sketch; solo esas
  definen extremos con pata L.
- El **segmento principal** de la barra es siempre el tramo distinto a las
  patas L (in-plane); dirección y extremos se toman de ese tramo.
- Shape: una pata → «02»; ambas patas → «03».
- Aplicación sobre ``Rebar`` libres (post RemoveAreaSystem), dentro de Tx abierta.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("System")

import System
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Curve,
    ElementId,
    FilteredElementCollector,
    Floor,
    Line,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    MultiplanarOption,
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarShape,
    RebarStyle,
)

try:
    from rebar_extender_l_ganchos_135_rps import (
        largo_pata_mm_desde_espesor_host,
        _try_create_l_from_rebar_shape_2seg,
    )
except Exception:
    largo_pata_mm_desde_espesor_host = None
    _try_create_l_from_rebar_shape_2seg = None

PATA_RESTA_ESPESOR_MM = 50.0
PATA_LARGO_MIN_MM = 10.0
PATA_LARGO_FALLBACK_MM = 150.0
# Tol. mm: extremo ↔ arista candidata (inset Create ~25 mm + recubrimiento AR)
PATA_EDGE_MATCH_TOL_MM = 220.0
PATA_HIT_TOL_MM = 120.0
# |dot(dir_barra, dir_arista)| < esto ⇒ extremos perpendiculares a la arista
PATA_PERP_DOT_MAX = 0.55
# Coincidencia arista paño ↔ outline/shaft/hueco
PATA_COINCIDE_PARALLEL_DOT = 0.92
PATA_COINCIDE_LATERAL_MM = 55.0
PATA_COINCIDE_MIN_OVERLAP_MM = 80.0
# Tol. para considerar un tramo de shaft/hueco “dentro” del paño
PATA_HOLE_IN_PANO_TOL_MM = 40.0
# |dot(dir_curva, dir_pata)| ≥ esto ⇒ tramo tipo pata (no principal)
PATA_SEG_ALIGN_DOT = 0.85
# RebarShape del proyecto (nombre visible)
PATA_SHAPE_ONE_END = u"02"
PATA_SHAPE_BOTH_ENDS = u"03"


def _as_unicode(val):
    try:
        return unicode(val)
    except Exception:
        try:
            return str(val)
        except Exception:
            return u""


def _mm_to_internal(mm):
    from Autodesk.Revit.DB import UnitTypeId, UnitUtils

    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def normalize_face_id(face_id):
    if face_id in (u"superior", u"inferior"):
        return face_id
    return u"inferior"


def pata_largo_mm_floor(document, floor):
    """Espesor losa − 50 mm (mín. ``PATA_LARGO_MIN_MM``)."""
    if largo_pata_mm_desde_espesor_host is not None and floor is not None:
        try:
            return float(
                largo_pata_mm_desde_espesor_host(
                    document,
                    floor,
                    resta_mm=PATA_RESTA_ESPESOR_MM,
                    fallback_mm=PATA_LARGO_FALLBACK_MM,
                    min_largo_mm=PATA_LARGO_MIN_MM,
                )
            )
        except Exception:
            pass
    return max(PATA_LARGO_MIN_MM, PATA_LARGO_FALLBACK_MM)


def pata_dir_xyz(plane, face_id):
    """
    Vector unitario de la pata L.

    Inferior → arriba (+normal con Z≥0); Superior → abajo (−normal).
    """
    face = normalize_face_id(face_id)
    n = None
    try:
        if plane is not None:
            n = plane.Normal
            if n is not None and float(n.GetLength()) > 1e-12:
                n = n.Normalize()
    except Exception:
        n = None
    if n is None:
        n = XYZ.BasisZ
    try:
        if float(n.Z) < 0.0:
            n = n.Negate()
    except Exception:
        pass
    if face == u"inferior":
        return n
    try:
        return n.Negate()
    except Exception:
        return XYZ(0, 0, -1)


def pano_edge_count(pts):
    if not pts or len(pts) < 3:
        return 0
    return len(pts)


def pano_edge_segment(pts, edge_idx):
    """((x0,y0), (x1,y1)) o None. Anillo cerrado implícito."""
    n = pano_edge_count(pts)
    if n < 3:
        return None
    i = int(edge_idx) % n
    j = (i + 1) % n
    try:
        return (
            (float(pts[i][0]), float(pts[i][1])),
            (float(pts[j][0]), float(pts[j][1])),
        )
    except Exception:
        return None


def _dist2_point_segment_mm(px, py, ax, ay, bx, by):
    """Distancia² punto→segmento en mm² + parámetro t en [0,1]."""
    abx = bx - ax
    aby = by - ay
    apx = px - ax
    apy = py - ay
    ab2 = abx * abx + aby * aby
    if ab2 < 1e-12:
        return apx * apx + apy * apy, 0.0
    t = (apx * abx + apy * aby) / ab2
    if t < 0.0:
        t = 0.0
    elif t > 1.0:
        t = 1.0
    qx = ax + abx * t
    qy = ay + aby * t
    dx = px - qx
    dy = py - qy
    return dx * dx + dy * dy, t


def hit_pano_edge_mm(pts, pt_mm, tol_mm=None):
    """Índice de arista más cercana a ``pt_mm``, o None si fuera de tol."""
    if tol_mm is None:
        tol_mm = PATA_HIT_TOL_MM
    if not pts or pt_mm is None or len(pts) < 3:
        return None
    try:
        px = float(pt_mm[0])
        py = float(pt_mm[1])
    except Exception:
        return None
    tol2 = float(tol_mm) * float(tol_mm)
    best_i = None
    best_d2 = None
    n = len(pts)
    for i in range(n):
        try:
            ax, ay = float(pts[i][0]), float(pts[i][1])
            bx, by = float(pts[(i + 1) % n][0]), float(pts[(i + 1) % n][1])
        except Exception:
            continue
        d2, _t = _dist2_point_segment_mm(px, py, ax, ay, bx, by)
        if d2 > tol2:
            continue
        if best_d2 is None or d2 < best_d2:
            best_d2 = d2
            best_i = i
    return best_i


def normalize_pata_edges(edges, n_pts):
    """Lista ordenada única de índices válidos."""
    out = []
    seen = set()
    n = int(n_pts or 0)
    if n < 3:
        return out
    for e in edges or []:
        try:
            i = int(e) % n
        except Exception:
            continue
        if i in seen:
            continue
        seen.add(i)
        out.append(i)
    out.sort()
    return out


def _xyz_to_plane_mm(pt, plane):
    if pt is None or plane is None:
        return None
    try:
        o = plane.Origin
        xv = plane.XVec
        yv = plane.YVec
        v = pt - o
        x_ft = float(v.DotProduct(xv))
        y_ft = float(v.DotProduct(yv))
        return (x_ft * 304.8, y_ft * 304.8)
    except Exception:
        return None


def _unit2(dx, dy):
    L = (float(dx) * float(dx) + float(dy) * float(dy)) ** 0.5
    if L < 1e-9:
        return None
    return (float(dx) / L, float(dy) / L)


def _edge_dir_mm(seg):
    if seg is None:
        return None
    (ax, ay), (bx, by) = seg
    return _unit2(bx - ax, by - ay)


def _pata_dir_ref_from_plane(plane):
    """Referencia unitaria de pata (+normal con Z≥0) para clasificar tramos."""
    n = None
    try:
        if plane is not None and plane.Normal is not None:
            n = plane.Normal.Normalize()
    except Exception:
        n = None
    if n is None:
        n = XYZ.BasisZ
    try:
        if float(n.Z) < 0.0:
            n = n.Negate()
    except Exception:
        pass
    return n


def _curve_length_ft(curve):
    try:
        return float(curve.Length)
    except Exception:
        pass
    try:
        return float(curve.GetEndPoint(0).DistanceTo(curve.GetEndPoint(1)))
    except Exception:
        return 0.0


def _curve_dir_xyz(curve):
    if curve is None:
        return None
    try:
        v = curve.GetEndPoint(1) - curve.GetEndPoint(0)
        if float(v.GetLength()) > 1e-12:
            return v.Normalize()
    except Exception:
        pass
    return _curve_tangent(curve, 0)


def _is_pata_like_curve(curve, pata_dir):
    """True si la curva es ~paralela a la dirección de pata L."""
    if curve is None or pata_dir is None:
        return False
    vd = _curve_dir_xyz(curve)
    if vd is None:
        return False
    try:
        pd = pata_dir.Normalize()
        return abs(float(vd.DotProduct(pd))) >= float(PATA_SEG_ALIGN_DOT)
    except Exception:
        return False


def _main_curve_from_chain(curves, pata_dir=None, plane=None):
    """
    Segmento principal: el tramo distinto a las patas L.

    Preferencia: curva más larga no alineada con ``pata_dir`` (±normal losa).
    """
    if not curves:
        return None
    if pata_dir is None:
        pata_dir = _pata_dir_ref_from_plane(plane)
    best = None
    best_len = -1.0
    best_any = None
    best_any_len = -1.0
    for c in curves:
        L = _curve_length_ft(c)
        if L > best_any_len:
            best_any_len = L
            best_any = c
        if _is_pata_like_curve(c, pata_dir):
            continue
        if L > best_len:
            best_len = L
            best = c
    return best if best is not None else best_any


def _main_curves_from_chain(curves, pata_dir=None, plane=None):
    """Curvas del tramo principal (todas las no-pata), o la principal sola."""
    if not curves:
        return []
    if pata_dir is None:
        pata_dir = _pata_dir_ref_from_plane(plane)
    body = [c for c in curves if not _is_pata_like_curve(c, pata_dir)]
    if body:
        return body
    main = _main_curve_from_chain(curves, pata_dir=pata_dir, plane=plane)
    return [main] if main is not None else list(curves)


def _bar_dir_mm_from_curves(curves, plane, pata_dir=None):
    """Dirección en plano-mm del tramo principal (no pata L)."""
    if not curves or plane is None:
        return None
    main = _main_curve_from_chain(curves, pata_dir=pata_dir, plane=plane)
    if main is None:
        return None
    try:
        p0 = main.GetEndPoint(0)
        p1 = main.GetEndPoint(1)
    except Exception:
        return None
    mm0 = _xyz_to_plane_mm(p0, plane)
    mm1 = _xyz_to_plane_mm(p1, plane)
    if mm0 is None or mm1 is None:
        return None
    return _unit2(mm1[0] - mm0[0], mm1[1] - mm0[1])


def _point_near_pata_edges(
    pt_mm, pts_ring, edge_indices, bar_dir_mm=None, tol_mm=None
):
    """
    True si ``pt_mm`` está cerca de alguna arista marcada.

    Si se pasa ``bar_dir_mm``, solo cuentan aristas ~perpendiculares a la barra
    (evita patas en barras paralelas a la arista marcada).
    """
    if pt_mm is None or not pts_ring or not edge_indices:
        return False
    segs = []
    for ei in edge_indices:
        seg = pano_edge_segment(pts_ring, ei)
        if seg is not None:
            segs.append(seg)
    return _point_near_segments(pt_mm, segs, bar_dir_mm=bar_dir_mm, tol_mm=tol_mm)


def _point_near_segments(pt_mm, segments, bar_dir_mm=None, tol_mm=None):
    """True si ``pt_mm`` está cerca de algún segmento (opcional ~perp. a barra)."""
    if pt_mm is None or not segments:
        return False
    if tol_mm is None:
        tol_mm = PATA_EDGE_MATCH_TOL_MM
    tol2 = float(tol_mm) * float(tol_mm)
    try:
        px, py = float(pt_mm[0]), float(pt_mm[1])
    except Exception:
        return False
    bdir = bar_dir_mm
    for seg in segments:
        if seg is None:
            continue
        if bdir is not None:
            edir = _edge_dir_mm(seg)
            if edir is None:
                continue
            try:
                dot = abs(
                    float(bdir[0]) * float(edir[0]) + float(bdir[1]) * float(edir[1])
                )
            except Exception:
                continue
            if dot > float(PATA_PERP_DOT_MAX):
                continue
        (ax, ay), (bx, by) = seg
        d2, _t = _dist2_point_segment_mm(px, py, ax, ay, bx, by)
        if d2 <= tol2:
            return True
    return False


def _ring_to_segments(pts_ring):
    """Lista de segmentos ((x0,y0),(x1,y1)) del anillo cerrado."""
    out = []
    n = pano_edge_count(pts_ring)
    for i in range(n):
        seg = pano_edge_segment(pts_ring, i)
        if seg is None:
            continue
        (ax, ay), (bx, by) = seg
        if (ax - bx) * (ax - bx) + (ay - by) * (ay - by) < 1e-6:
            continue
        out.append(seg)
    return out


def _point_in_ring_mm(pt_mm, ring, tol_mm=None):
    """Ray-cast; borde cuenta como dentro (tol. a aristas)."""
    if pt_mm is None or not ring or len(ring) < 3:
        return False
    try:
        px, py = float(pt_mm[0]), float(pt_mm[1])
    except Exception:
        return False
    if tol_mm is None:
        tol_mm = PATA_HOLE_IN_PANO_TOL_MM
    tol2 = float(tol_mm) * float(tol_mm)
    # Cerca de arista → dentro
    for seg in _ring_to_segments(ring):
        (ax, ay), (bx, by) = seg
        d2, _t = _dist2_point_segment_mm(px, py, ax, ay, bx, by)
        if d2 <= tol2:
            return True
    inside = False
    n = len(ring)
    for i in range(n):
        try:
            x1, y1 = float(ring[i][0]), float(ring[i][1])
            x2, y2 = float(ring[(i + 1) % n][0]), float(ring[(i + 1) % n][1])
        except Exception:
            continue
        cond = (y1 > py) != (y2 > py)
        if not cond:
            continue
        try:
            x_int = (x2 - x1) * (py - y1) / (y2 - y1 + 1e-30) + x1
        except Exception:
            continue
        if px < x_int:
            inside = not inside
    return inside


def _segments_coincide_mm(seg_a, seg_b):
    """
    True si dos segmentos son ~colineales, cercanos lateralmente y se solapan.
    """
    if seg_a is None or seg_b is None:
        return False
    da = _edge_dir_mm(seg_a)
    db = _edge_dir_mm(seg_b)
    if da is None or db is None:
        return False
    try:
        dot = abs(float(da[0]) * float(db[0]) + float(da[1]) * float(db[1]))
    except Exception:
        return False
    if dot < float(PATA_COINCIDE_PARALLEL_DOT):
        return False
    (ax, ay), (bx, by) = seg_a
    (cx, cy), (dx, dy) = seg_b
    lat = float(PATA_COINCIDE_LATERAL_MM)
    lat2 = lat * lat

    def _near_seg(px, py, s0, s1):
        d2, _t = _dist2_point_segment_mm(px, py, s0[0], s0[1], s1[0], s1[1])
        return d2 <= lat2

    hits_a = _near_seg(ax, ay, (cx, cy), (dx, dy)) or _near_seg(
        bx, by, (cx, cy), (dx, dy)
    )
    hits_b = _near_seg(cx, cy, (ax, ay), (bx, by)) or _near_seg(
        dx, dy, (ax, ay), (bx, by)
    )
    mx_a, my_a = 0.5 * (ax + bx), 0.5 * (ay + by)
    mx_b, my_b = 0.5 * (cx + dx), 0.5 * (cy + dy)
    hits_mid = _near_seg(mx_a, my_a, (cx, cy), (dx, dy)) or _near_seg(
        mx_b, my_b, (ax, ay), (bx, by)
    )
    if not (hits_a or hits_b or hits_mid):
        return False
    # Solape proyectado sobre eje de A
    try:
        abx, aby = bx - ax, by - ay
        lab = (abx * abx + aby * aby) ** 0.5
        if lab < 1e-9:
            return False
        ux, uy = abx / lab, aby / lab

        def _proj(px, py):
            return (px - ax) * ux + (py - ay) * uy

        t0 = min(_proj(cx, cy), _proj(dx, dy))
        t1 = max(_proj(cx, cy), _proj(dx, dy))
        o0 = max(0.0, t0)
        o1 = min(lab, t1)
        overlap = o1 - o0
    except Exception:
        return False
    return overlap >= float(PATA_COINCIDE_MIN_OVERLAP_MM)


def _seg_related_to_pano(seg, pano_pts):
    """True si el segmento corta / toca / está dentro del paño."""
    if seg is None or not pano_pts or len(pano_pts) < 3:
        return False
    (ax, ay), (bx, by) = seg
    # Extremos o punto medio dentro / cerca del paño
    for pt in ((ax, ay), (bx, by), (0.5 * (ax + bx), 0.5 * (ay + by))):
        if _point_in_ring_mm(pt, pano_pts):
            return True
    # Coincide con alguna arista del paño
    for pseg in _ring_to_segments(pano_pts):
        if _segments_coincide_mm(seg, pseg):
            return True
    return False


def select_pata_candidate_segments(pano_pts, outline_pts, hole_rings=None):
    """
    Aristas que definen pata L para un paño.

    1) Aristas del paño que coinciden con outline o con shafts/huecos.
    2) Aristas de shafts/huecos que interactúan con el paño (cortes).
    """
    pano = _normalize_rings_mm([pano_pts])
    if not pano:
        return []
    pano_pts = pano[0]
    pano_segs = _ring_to_segments(pano_pts)
    ref_segs = []
    for ring in _normalize_rings_mm(
        ([outline_pts] if outline_pts else []) + list(hole_rings or [])
    ):
        ref_segs.extend(_ring_to_segments(ring))
    if not ref_segs:
        return []
    candidates = []
    seen = set()

    def _add(seg):
        if seg is None:
            return
        (ax, ay), (bx, by) = seg
        key = (
            round(min(ax, bx), 1),
            round(min(ay, by), 1),
            round(max(ax, bx), 1),
            round(max(ay, by), 1),
        )
        if key in seen:
            return
        seen.add(key)
        candidates.append(seg)

    # Perímetro del paño que coincide con referencia (outline / shaft / hueco)
    for pseg in pano_segs:
        for rseg in ref_segs:
            if _segments_coincide_mm(pseg, rseg):
                _add(pseg)
                break
    # Cortes: aristas de shaft/hueco (no outline exterior completo) que tocan el paño
    hole_only = []
    for ring in _normalize_rings_mm(hole_rings or []):
        hole_only.extend(_ring_to_segments(ring))
    for hseg in hole_only:
        if _seg_related_to_pano(hseg, pano_pts):
            _add(hseg)
    return candidates


def _centerline_curves_rebar(rebar, pos_idx=0):
    if rebar is None:
        return []
    for mpo_name in (
        u"IncludeAllMultiplanarCurves",
        u"IncludeOnlyPlanarCurves",
    ):
        mpo = getattr(MultiplanarOption, mpo_name, None)
        if mpo is None:
            continue
        try:
            raw = rebar.GetCenterlineCurves(False, False, False, mpo, int(pos_idx))
            if raw is not None and int(raw.Count) > 0:
                return [raw[i] for i in range(int(raw.Count))]
        except Exception:
            pass
    try:
        raw = rebar.GetCenterlineCurves(False, False, False)
        if raw is not None and int(raw.Count) > 0:
            return [raw[i] for i in range(int(raw.Count))]
    except Exception:
        pass
    return []


def _rebar_normal(rebar, plane=None):
    try:
        n = rebar.GetShapeDrivenAccessor().Normal
        if n is not None and float(n.GetLength()) > 1e-12:
            return n.Normalize()
    except Exception:
        pass
    try:
        if plane is not None and plane.Normal is not None:
            return plane.Normal.Normalize()
    except Exception:
        pass
    return XYZ.BasisZ


def _curve_tangent(curve, at_end=0):
    """Tangente unitaria en extremo 0 o 1."""
    if curve is None:
        return None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        if int(at_end) == 0:
            v = p1 - p0
        else:
            v = p1 - p0
        if float(v.GetLength()) > 1e-12:
            return v.Normalize()
    except Exception:
        pass
    try:
        der = curve.ComputeDerivatives(float(at_end), True)
        v = der.BasisX
        if v is not None and float(v.GetLength()) > 1e-12:
            return v.Normalize()
    except Exception:
        pass
    return None


def _normals_for_vertical_l(curves, pata_dir, plane=None, rebar=None):
    """
    Normales candidatas para CreateFromCurves con pata vertical.

    La L (barra horizontal + pata ±Z) vive en un plano vertical: la normal debe
    ser horizontal (= tangente × dir_pata). Usar la normal del AR (≈Z) falla
    porque la pata es paralela a esa normal.
    """
    out = []
    seen = set()

    def _add(n):
        if n is None:
            return
        try:
            nn = n.Normalize()
        except Exception:
            return
        if float(nn.GetLength()) < 1e-12:
            return
        try:
            key = (
                round(float(nn.X), 5),
                round(float(nn.Y), 5),
                round(float(nn.Z), 5),
            )
        except Exception:
            return
        if key in seen:
            return
        seen.add(key)
        out.append(nn)
        try:
            _add(nn.Negate())
        except Exception:
            pass

    t = None
    if curves:
        t = _curve_tangent(curves[0], 0)
        if t is None and len(curves) > 0:
            try:
                t = (curves[0].GetEndPoint(1) - curves[0].GetEndPoint(0)).Normalize()
            except Exception:
                t = None
    d = None
    try:
        if pata_dir is not None:
            d = pata_dir.Normalize()
    except Exception:
        d = pata_dir
    if t is not None and d is not None:
        try:
            _add(t.CrossProduct(d))
        except Exception:
            pass
        try:
            _add(d.CrossProduct(t))
        except Exception:
            pass
    if rebar is not None:
        _add(_rebar_normal(rebar, plane))
    if plane is not None:
        try:
            _add(plane.Normal)
        except Exception:
            pass
    if not out:
        _add(XYZ.BasisY)
        _add(XYZ.BasisX)
    return out


def _hook_orient(rebar, end):
    try:
        o = rebar.GetHookOrientation(int(end))
        if o is not None:
            return o
    except Exception:
        pass
    return RebarHookOrientation.Left


def _copy_layout_shape_driven(src, dst):
    try:
        a0 = src.GetShapeDrivenAccessor()
        a1 = dst.GetShapeDrivenAccessor()
    except Exception:
        return False
    if a0 is None or a1 is None:
        return False
    rule_name = u""
    try:
        rule_name = src.LayoutRule.ToString()
    except Exception:
        try:
            rule_name = a0.GetLayoutRule().ToString()
        except Exception:
            rule_name = u""
    try:
        sp = float(src.MaxSpacing)
    except Exception:
        sp = 0.0
    try:
        alen = float(a0.ArrayLength)
    except Exception:
        try:
            alen = float(a0.GetArrayLength())
        except Exception:
            alen = 0.0
    try:
        b_side = bool(a0.BarsOnNormalSide)
    except Exception:
        b_side = True
    try:
        inc0 = bool(src.IncludeFirstBar)
        inc1 = bool(src.IncludeLastBar)
    except Exception:
        inc0 = inc1 = True
    try:
        nbars = int(src.Quantity)
    except Exception:
        nbars = 1
    try:
        if rule_name == u"Single":
            a1.SetLayoutAsSingle()
        elif rule_name == u"MaximumSpacing":
            a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
        elif rule_name in (u"Number", u"FixedNumber"):
            a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
        elif rule_name == u"NumberWithSpacing":
            a1.SetLayoutAsNumberWithSpacing(nbars, sp, alen, b_side, inc0, inc1)
        elif rule_name == u"MinimumClearSpacing":
            a1.SetLayoutAsMinimumClearSpacing(sp, alen, b_side, inc0, inc1)
        else:
            a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
        return True
    except Exception:
        return False


def _rebar_shape_visible_name(shape):
    if shape is None:
        return u""
    for bip in (
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
    ):
        try:
            p = shape.get_Parameter(bip)
            if p is not None and p.HasValue:
                s = (p.AsString() or u"").strip()
                if s:
                    return s
        except Exception:
            continue
    try:
        return (getattr(shape, u"Name", None) or u"").strip()
    except Exception:
        return u""


def _find_rebar_shape_by_name(document, nombre):
    """``RebarShape`` por nombre visible (exacto / case / dígitos «02»)."""
    if document is None or not nombre:
        return None
    key = (nombre or u"").strip()
    if not key:
        return None
    try:
        key_low = key.lower()
    except Exception:
        key_low = key
    key_digits = u"".join(ch for ch in key if ch in u"0123456789")
    candidates = []
    try:
        for sh in FilteredElementCollector(document).OfClass(RebarShape):
            sn = _rebar_shape_visible_name(sh)
            if not sn:
                continue
            try:
                sn_low = sn.lower()
            except Exception:
                sn_low = sn
            dig = u"".join(ch for ch in sn if ch in u"0123456789")
            candidates.append((sh, sn, sn_low, dig))
    except Exception:
        return None
    for sh, sn, _sl, _d in candidates:
        if sn == key:
            return sh
    for sh, _sn, sn_low, _d in candidates:
        if sn_low == key_low:
            return sh
    if key_digits:
        for sh, _sn, _sl, dig in candidates:
            if dig == key_digits:
                return sh
    return None


def _apply_rebar_style_if_writable(rebar, style):
    if rebar is None or style is None:
        return
    try:
        rebar.Style = style
    except Exception:
        pass


def _try_create_from_shape_named(
    document, curves_list, host, norm, bar_type, style, o0, o1, shape_name
):
    """``CreateFromCurvesAndShape`` con shape del proyecto (p. ej. «02» / «03»)."""
    if document is None or not curves_list or host is None or bar_type is None:
        return None
    shape = _find_rebar_shape_by_name(document, shape_name)
    if shape is None:
        return None
    try:
        cl = List[Curve]()
        for c in curves_list:
            cl.Add(c)
    except Exception:
        return None
    orient_tries = (
        (o0, o1),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    )
    seen = set()
    pairs = []
    for a in orient_tries:
        try:
            k = (int(a[0]), int(a[1]))
        except Exception:
            k = (str(a[0]), str(a[1]))
        if k in seen:
            continue
        seen.add(k)
        pairs.append(a)
    inv = ElementId.InvalidElementId
    for so, eo in pairs:
        try:
            r = Rebar.CreateFromCurvesAndShape(
                document,
                shape,
                bar_type,
                None,
                None,
                host,
                norm,
                cl,
                so,
                eo,
                0.0,
                0.0,
                inv,
                inv,
            )
            if r is not None:
                _apply_rebar_style_if_writable(r, style)
                return r
        except Exception:
            pass
        try:
            r = Rebar.CreateFromCurvesAndShape(
                document, shape, bar_type, None, None, host, norm, cl, so, eo
            )
            if r is not None:
                _apply_rebar_style_if_writable(r, style)
                return r
        except Exception:
            pass
    return None


def _try_assign_rebar_shape_name(document, rebar, shape_name):
    """Asigna shape por nombre tras Create (respaldo)."""
    if document is None or rebar is None or not shape_name:
        return False
    shape = _find_rebar_shape_by_name(document, shape_name)
    if shape is None:
        return False
    try:
        sid = shape.Id
        if sid is None or sid == ElementId.InvalidElementId:
            return False
    except Exception:
        return False
    for meth_name in (u"ChangeTypeId", u"ChangeTypeId"):
        try:
            meth = getattr(rebar, meth_name, None)
            if meth is not None:
                meth(sid)
                return True
        except Exception:
            pass
    try:
        # Algunas builds: parámetro de tipo de shape
        p = rebar.get_Parameter(BuiltInParameter.REBAR_SHAPE)
        if p is not None and (not p.IsReadOnly):
            p.Set(sid)
            return True
    except Exception:
        pass
    return False


def _create_from_curves_no_hooks(doc, curves_list, host, norm, bar_type, style, o0, o1):
    """Prueba List/Array y combinaciones useExisting/createNew."""
    if doc is None or not curves_list or host is None or bar_type is None or norm is None:
        return None
    if style is None:
        style = RebarStyle.Standard
    arr = None
    cl = None
    try:
        cl = List[Curve]()
        for c in curves_list:
            cl.Add(c)
    except Exception:
        cl = None
    try:
        ct = clr.GetClrType(Line).BaseType
        n = len(curves_list)
        arr = System.Array.CreateInstance(ct, n)
        for i in range(n):
            arr[i] = curves_list[i]
    except Exception:
        arr = None
    orients = (
        (o0, o1),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
    )
    seen_o = set()
    for so, eo in orients:
        try:
            ok = (int(so), int(eo))
        except Exception:
            ok = (str(so), str(eo))
        if ok in seen_o:
            continue
        seen_o.add(ok)
        for curves_arg in (cl, arr):
            if curves_arg is None:
                continue
            for use_ex, create_new in (
                (True, True),
                (False, True),
                (True, False),
                (False, False),
            ):
                try:
                    r = Rebar.CreateFromCurves(
                        doc,
                        style,
                        bar_type,
                        None,
                        None,
                        host,
                        norm,
                        curves_arg,
                        so,
                        eo,
                        use_ex,
                        create_new,
                    )
                    if r is not None:
                        return r
                except Exception:
                    continue
    return None


def _normalize_rings_mm(rings):
    """Lista de anillos válidos (≥3 pts) como tuples (x,y) float."""
    out = []
    for ring in rings or []:
        pts = []
        for p in ring or []:
            try:
                pts.append((float(p[0]), float(p[1])))
            except Exception:
                continue
        if len(pts) >= 3:
            out.append(pts)
    return out


def _ends_need_pata(rebar, plane, pts_ring, edge_indices, tol_mm=None):
    """(pata_start, pata_end) si extremos tocan aristas ~perpendiculares a la barra."""
    return _ends_need_pata_on_rings(
        rebar, plane, [pts_ring], edge_indices_per_ring=[edge_indices], tol_mm=tol_mm
    )


def _ends_need_pata_on_rings(
    rebar, plane, rings, edge_indices_per_ring=None, tol_mm=None
):
    """
    (pata_start, pata_end) si algún extremo toca aristas de cualquiera de ``rings``.

    Si ``edge_indices_per_ring`` es None, usa todas las aristas de cada anillo.
    """
    segs = []
    rings_n = _normalize_rings_mm(rings)
    for i, ring in enumerate(rings_n):
        if edge_indices_per_ring is not None and i < len(edge_indices_per_ring):
            edges = normalize_pata_edges(
                edge_indices_per_ring[i], pano_edge_count(ring)
            )
            for ei in edges:
                seg = pano_edge_segment(ring, ei)
                if seg is not None:
                    segs.append(seg)
        else:
            segs.extend(_ring_to_segments(ring))
    return _ends_need_pata_on_segments(rebar, plane, segs, tol_mm=tol_mm)


def _ends_need_pata_on_segments(rebar, plane, segments, tol_mm=None):
    """
    (pata_start, pata_end) si extremos del **tramo principal** tocan candidatos.

    El segmento principal es el distinto a las patas L (±normal); no se usan
    las puntas de patas ya existentes.
    """
    curves = _centerline_curves_rebar(rebar, 0)
    if not curves or not segments:
        return False, False
    pata_ref = _pata_dir_ref_from_plane(plane)
    body = _main_curves_from_chain(curves, pata_dir=pata_ref, plane=plane)
    if not body:
        return False, False
    try:
        p0 = body[0].GetEndPoint(0)
        p1 = body[-1].GetEndPoint(1)
    except Exception:
        return False, False
    mm0 = _xyz_to_plane_mm(p0, plane)
    mm1 = _xyz_to_plane_mm(p1, plane)
    bdir = _bar_dir_mm_from_curves(curves, plane, pata_dir=pata_ref)
    return (
        _point_near_segments(mm0, segments, bar_dir_mm=bdir, tol_mm=tol_mm),
        _point_near_segments(mm1, segments, bar_dir_mm=bdir, tol_mm=tol_mm),
    )


def extend_rebar_pata_l(
    document,
    rebar,
    pata_start,
    pata_end,
    dir_xyz,
    largo_mm,
    plane=None,
    avisos=None,
):
    """
    Sustituye ``rebar`` por uno con patas L geométricas en los extremos pedidos.

    Requiere Transaction abierta. Devuelve el nuevo ``Rebar`` o el original si falla.
    """
    if avisos is None:
        avisos = []
    if document is None or rebar is None or not (pata_start or pata_end):
        return rebar
    if dir_xyz is None or float(largo_mm or 0.0) < PATA_LARGO_MIN_MM:
        return rebar
    try:
        rid = int(rebar.Id.IntegerValue)
    except Exception:
        rid = 0
    try:
        host = document.GetElement(rebar.GetHostId())
    except Exception:
        host = None
    if host is None:
        avisos.append(u"Pata L Id {0}: sin host.".format(rid))
        return rebar
    curves = _centerline_curves_rebar(rebar, 0)
    if not curves:
        avisos.append(u"Pata L Id {0}: sin centerline.".format(rid))
        return rebar
    try:
        d = dir_xyz.Normalize()
        le = _mm_to_internal(largo_mm)
    except Exception as ex:
        avisos.append(u"Pata L Id {0}: dir/largo ({1}).".format(rid, _as_unicode(ex)))
        return rebar
    # Cuerpo = segmento(s) principal(es), distintos a patas L
    body = _main_curves_from_chain(curves, pata_dir=d, plane=plane)
    if not body:
        avisos.append(u"Pata L Id {0}: sin tramo principal.".format(rid))
        return rebar
    new_chain = list(body)
    try:
        if pata_start:
            p0 = new_chain[0].GetEndPoint(0)
            tip = p0 + d.Multiply(le)
            new_chain = [Line.CreateBound(tip, p0)] + new_chain
        if pata_end:
            p1 = new_chain[-1].GetEndPoint(1)
            tip = p1 + d.Multiply(le)
            new_chain = new_chain + [Line.CreateBound(p1, tip)]
    except Exception as ex:
        avisos.append(u"Pata L Id {0}: geometría ({1}).".format(rid, _as_unicode(ex)))
        return rebar
    try:
        bar_type = document.GetElement(rebar.GetTypeId())
    except Exception:
        bar_type = None
    if not isinstance(bar_type, RebarBarType):
        avisos.append(u"Pata L Id {0}: sin RebarBarType.".format(rid))
        return rebar
    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    o0 = _hook_orient(rebar, 0)
    o1 = _hook_orient(rebar, 1)
    orig_id = rebar.Id
    # Shape: una pata → 02; ambas → 03
    shape_name = (
        PATA_SHAPE_BOTH_ENDS
        if (pata_start and pata_end)
        else PATA_SHAPE_ONE_END
    )
    # Normal = tangente del tramo principal × dir_pata
    norms = _normals_for_vertical_l(body, d, plane=plane, rebar=rebar)
    new_rb = None
    last_err = u""
    used_named_shape = False
    for norm in norms:
        try:
            new_rb = _try_create_from_shape_named(
                document,
                new_chain,
                host,
                norm,
                bar_type,
                style,
                o0,
                o1,
                shape_name,
            )
        except Exception as ex_nm:
            new_rb = None
            last_err = _as_unicode(ex_nm)
        if new_rb is not None:
            used_named_shape = True
            break
        if _try_create_l_from_rebar_shape_2seg is not None:
            try:
                new_rb = _try_create_l_from_rebar_shape_2seg(
                    document, new_chain, host, norm, bar_type, style, o0, o1
                )
            except Exception as ex_sh:
                new_rb = None
                last_err = _as_unicode(ex_sh)
        if new_rb is None:
            try:
                new_rb = _create_from_curves_no_hooks(
                    document, new_chain, host, norm, bar_type, style, o0, o1
                )
            except Exception as ex_cf:
                new_rb = None
                last_err = _as_unicode(ex_cf)
        if new_rb is not None:
            break
    if new_rb is None:
        avisos.append(
            u"Pata L Id {0}: Create falló (shape {1}){2}.".format(
                rid,
                shape_name,
                (u" ({0})".format(last_err) if last_err else u""),
            )
        )
        return rebar
    if not used_named_shape:
        if not _try_assign_rebar_shape_name(document, new_rb, shape_name):
            avisos.append(
                u"Pata L Id {0}: no se asignó shape {1} "
                u"(creada con forma alternativa).".format(rid, shape_name)
            )
    if not _copy_layout_shape_driven(rebar, new_rb):
        avisos.append(
            u"Pata L Id {0}: layout no copiado (se mantiene nuevo).".format(rid)
        )
    try:
        if orig_id is not None and orig_id != ElementId.InvalidElementId:
            document.Delete(orig_id)
    except Exception as ex:
        avisos.append(
            u"Pata L Id {0}: no se eliminó original ({1}).".format(
                rid, _as_unicode(ex)
            )
        )
    return new_rb


def _aplicar_patas_l_on_segments(
    document,
    rebars,
    floor,
    plane,
    face_id,
    segments,
    avisos=None,
):
    """Núcleo: patas L si extremos tocan ``segments`` candidatos."""
    if avisos is None:
        avisos = []
    if not segments or not rebars:
        if not segments:
            avisos.append(
                u"Pata L: ninguna arista del paño coincide con outline/shaft/hueco."
            )
        return list(rebars or [])
    if plane is None:
        avisos.append(u"Pata L: sin plano de Sketch; no se aplicaron.")
        return list(rebars or [])
    largo = pata_largo_mm_floor(document, floor)
    direction = pata_dir_xyz(plane, face_id)
    if direction is None:
        avisos.append(u"Pata L: sin dirección de pata; no se aplicaron.")
        return list(rebars or [])
    out = []
    n_cand = 0
    n_ok = 0
    for rb in rebars or []:
        try:
            rb = document.GetElement(rb.Id)
        except Exception:
            pass
        if not isinstance(rb, Rebar):
            if rb is not None:
                out.append(rb)
            continue
        p_start, p_end = _ends_need_pata_on_segments(rb, plane, segments)
        if not p_start and not p_end:
            out.append(rb)
            continue
        n_cand += 1
        new_rb = extend_rebar_pata_l(
            document,
            rb,
            p_start,
            p_end,
            direction,
            largo,
            plane=plane,
            avisos=avisos,
        )
        if new_rb is not None and new_rb is not rb:
            try:
                if int(new_rb.Id.IntegerValue) != int(rb.Id.IntegerValue):
                    n_ok += 1
            except Exception:
                n_ok += 1
        out.append(new_rb if new_rb is not None else rb)
    if n_cand <= 0:
        avisos.append(
            u"Pata L: {0} arista(s) candidata(s), ningún extremo de barra "
            u"coincidió (~perpendicular).".format(len(segments))
        )
    elif n_ok <= 0:
        avisos.append(
            u"Pata L: {0} barra(s) candidata(s), Create falló en todas "
            u"(largo ≈ {1:g} mm).".format(n_cand, float(largo))
        )
    else:
        avisos.append(
            u"Pata L: {0}/{1} barra(s) con L (largo ≈ {2:g} mm, {3} arista(s)).".format(
                n_ok, n_cand, float(largo), len(segments)
            )
        )
    return out


def _aplicar_patas_l_on_rings(
    document,
    rebars,
    floor,
    plane,
    face_id,
    rings,
    avisos=None,
):
    """Compat: todas las aristas de ``rings`` como candidatas."""
    segs = []
    for ring in _normalize_rings_mm(rings):
        segs.extend(_ring_to_segments(ring))
    return _aplicar_patas_l_on_segments(
        document, rebars, floor, plane, face_id, segs, avisos=avisos
    )


def aplicar_patas_l_a_rebars(
    document,
    rebars,
    floor,
    plane,
    face_id,
    pts_ring,
    edge_indices,
    avisos=None,
):
    """
    Aplica patas L a rebars cuyos extremos tocan aristas indicadas de un anillo.
    """
    if avisos is None:
        avisos = []
    edges = normalize_pata_edges(edge_indices, pano_edge_count(pts_ring))
    if not edges or not rebars:
        return list(rebars or [])
    rings_n = _normalize_rings_mm([pts_ring])
    if not rings_n:
        return list(rebars or [])
    if plane is None:
        avisos.append(u"Pata L: sin plano de Sketch; no se aplicaron.")
        return list(rebars or [])
    largo = pata_largo_mm_floor(document, floor)
    direction = pata_dir_xyz(plane, face_id)
    if direction is None:
        avisos.append(u"Pata L: sin dirección de pata; no se aplicaron.")
        return list(rebars or [])
    out = []
    n_cand = 0
    n_ok = 0
    for rb in rebars or []:
        try:
            rb = document.GetElement(rb.Id)
        except Exception:
            pass
        if not isinstance(rb, Rebar):
            if rb is not None:
                out.append(rb)
            continue
        p_start, p_end = _ends_need_pata_on_rings(
            rb, plane, rings_n, edge_indices_per_ring=[edges]
        )
        if not p_start and not p_end:
            out.append(rb)
            continue
        n_cand += 1
        new_rb = extend_rebar_pata_l(
            document,
            rb,
            p_start,
            p_end,
            direction,
            largo,
            plane=plane,
            avisos=avisos,
        )
        if new_rb is not None and new_rb is not rb:
            try:
                if int(new_rb.Id.IntegerValue) != int(rb.Id.IntegerValue):
                    n_ok += 1
            except Exception:
                n_ok += 1
        out.append(new_rb if new_rb is not None else rb)
    return out


def aplicar_patas_l_por_outline(
    document,
    rebars,
    floor,
    plane,
    face_id,
    outline_pts,
    avisos=None,
    hole_rings=None,
    pano_pts=None,
):
    """
    Aplica patas L según aristas del paño que coinciden con outline/shaft/hueco.

    1. Compara aristas del ``pano_pts`` con outline + shafts/huecos.
    2. Añade aristas de shaft/hueco que cortan el paño.
    3. Solo esas aristas candidatas definen extremos con pata L.
    """
    if avisos is None:
        avisos = []
    if not pano_pts or len(pano_pts) < 3:
        avisos.append(u"Pata L: paño sin polígono; no se aplicaron.")
        return list(rebars or [])
    if not (outline_pts or hole_rings):
        avisos.append(u"Pata L: sin outline/huecos; no se aplicaron.")
        return list(rebars or [])
    segs = select_pata_candidate_segments(pano_pts, outline_pts, hole_rings)
    return _aplicar_patas_l_on_segments(
        document,
        rebars,
        floor,
        plane,
        face_id,
        segs,
        avisos=avisos,
    )
