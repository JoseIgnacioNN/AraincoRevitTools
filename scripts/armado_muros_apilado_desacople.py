# -*- coding: utf-8 -*-
"""
Apilamiento con largos distintos — zonas solape / desacople.

Horizontales: un solo AR por muro (continuidad).
Verticales: en post se parten sets que cruzan solape+desacople;
solape → empotramiento; desacople → pata L.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    Curve,
    Line,
    LocationCurve,
    Transform,
    UnitUtils,
    UnitTypeId,
    Wall,
    XYZ,
)

# Longitud mínima de paño (mm). Tramos menores se fusionan al vecino.
_MIN_PANO_MM = 250.0

# False = comportamiento previo a desacople (no partir verticales; cabeza
# usa apilado-sobre clásico). True = solape/desacople (split DividirRebarSet).
ENABLE_DESACOPLE = True

_LAST_CREATE_FAIL = u""


def log_desacople(msg, sink=None, to_ui=True):
    """Stub: instrumentación de desacople desactivada."""
    return u""


def _fmt_intervals(intervals):
    parts = []
    for a, b in intervals or []:
        try:
            parts.append(u"[{0:.3f},{1:.3f}]".format(float(a), float(b)))
        except Exception:
            parts.append(u"?")
    return u",".join(parts) if parts else u"(ninguno)"


def _fmt_idx_preview(idxs, limit=12):
    xs = list(idxs or [])
    if not xs:
        return u"[]"
    if len(xs) <= limit:
        return unicode(xs)
    head = u",".join(unicode(i) for i in xs[:limit])
    return u"[{0}… n={1}]".format(head, len(xs))


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _element_id_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _location_curve(wall):
    if wall is None or not isinstance(wall, Wall):
        return None
    try:
        loc = wall.Location
    except Exception:
        return None
    if not isinstance(loc, LocationCurve):
        return None
    try:
        return loc.Curve
    except Exception:
        return None


def _tangent_xy(curve):
    if curve is None:
        return None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        v = XYZ(float(p1.X) - float(p0.X), float(p1.Y) - float(p0.Y), 0.0)
        if float(v.GetLength()) < 1e-12:
            return None
        return v.Normalize()
    except Exception:
        return None


def _u_along_host(host_curve, pt, tang_xy=None):
    """Parámetro normalizado 0..1 de ``pt`` proyectado sobre el eje del host."""
    if host_curve is None or pt is None:
        return None
    try:
        p0 = host_curve.GetEndPoint(0)
        p1 = host_curve.GetEndPoint(1)
    except Exception:
        return None
    try:
        L = float(host_curve.Length)
    except Exception:
        L = 0.0
    if L < 1e-12:
        return None
    t = tang_xy
    if t is None:
        t = _tangent_xy(host_curve)
    if t is None:
        return None
    try:
        vx = float(pt.X) - float(p0.X)
        vy = float(pt.Y) - float(p0.Y)
        u = (vx * float(t.X) + vy * float(t.Y)) / L
    except Exception:
        return None
    if u < 0.0:
        return 0.0
    if u > 1.0:
        return 1.0
    return float(u)


def _intervalos_solape_u(host, walls_sel, muro_apilado_sobre_fn):
    """
    Intervalos [u0,u1] ⊂ [0,1] del host cubiertos por muros apilados encima.

    Solo ``LocationCurve`` del muro de encima (extremos + medio), con un pad
    pequeño (~50 mm). **No** usa esquinas de bbox (inflaban el solape y
    ocultaban el desacople).
    """
    lc = _location_curve(host)
    if lc is None:
        return []
    tang = _tangent_xy(lc)
    if tang is None:
        return []
    try:
        L = float(lc.Length)
    except Exception:
        L = 0.0
    if L < 1e-12:
        return []
    try:
        pad_u = float(_mm_to_internal(50.0)) / L
    except Exception:
        pad_u = 0.02
    hid = _element_id_int(getattr(host, "Id", None))
    intervals = []
    for other in walls_sel or []:
        if other is None or not isinstance(other, Wall):
            continue
        oid = _element_id_int(getattr(other, "Id", None))
        if hid is not None and oid is not None and int(hid) == int(oid):
            continue
        try:
            if not muro_apilado_sobre_fn(host, other):
                continue
        except Exception:
            continue
        oc = _location_curve(other)
        if oc is None:
            continue
        pts = []
        try:
            pts.append(oc.GetEndPoint(0))
            pts.append(oc.GetEndPoint(1))
            pts.append(oc.Evaluate(0.5, True))
        except Exception:
            continue
        us = []
        for pt in pts:
            u = _u_along_host(lc, pt, tang)
            if u is not None:
                us.append(u)
        if not us:
            continue
        u0 = max(0.0, min(us) - pad_u)
        u1 = min(1.0, max(us) + pad_u)
        if u1 - u0 < 1e-6:
            continue
        intervals.append((u0, u1))
    return _merge_intervals(intervals)


def _merge_intervals(intervals):
    if not intervals:
        return []
    ordered = sorted(intervals, key=lambda it: (it[0], it[1]))
    out = [list(ordered[0])]
    for a, b in ordered[1:]:
        if a <= out[-1][1] + 1e-6:
            out[-1][1] = max(out[-1][1], b)
        else:
            out.append([a, b])
    return [(float(a), float(b)) for a, b in out]


def _complement_intervals(covered, u0=0.0, u1=1.0):
    """Complemento de ``covered`` dentro de [u0,u1]."""
    if not covered:
        return [(float(u0), float(u1))]
    out = []
    cur = float(u0)
    for a, b in covered:
        if a > cur + 1e-9:
            out.append((cur, float(a)))
        cur = max(cur, float(b))
    if cur < float(u1) - 1e-9:
        out.append((cur, float(u1)))
    return out


def _min_pano_u(host):
    lc = _location_curve(host)
    try:
        L = float(lc.Length) if lc is not None else 0.0
    except Exception:
        L = 0.0
    if L < 1e-9:
        return 0.05
    try:
        min_ft = float(_mm_to_internal(_MIN_PANO_MM))
    except Exception:
        min_ft = 0.8
    return max(1e-4, min_ft / L)


def _fuse_small_zones(zones, min_u):
    """Fusiona tramos más cortos que ``min_u`` con el vecino (prioridad solape)."""
    if not zones:
        return []
    items = [dict(z) for z in zones]
    changed = True
    while changed and len(items) > 1:
        changed = False
        for i, z in enumerate(items):
            if float(z[u"u1"]) - float(z[u"u0"]) + 1e-12 >= float(min_u):
                continue
            # Fusionar con vecino: preferir solape si existe
            left = items[i - 1] if i > 0 else None
            right = items[i + 1] if i + 1 < len(items) else None
            target = None
            if left is not None and right is not None:
                if left[u"kind"] == u"solape":
                    target = left
                elif right[u"kind"] == u"solape":
                    target = right
                else:
                    target = left
            elif left is not None:
                target = left
            elif right is not None:
                target = right
            if target is None:
                continue
            target[u"u0"] = min(float(target[u"u0"]), float(z[u"u0"]))
            target[u"u1"] = max(float(target[u"u1"]), float(z[u"u1"]))
            items.pop(i)
            changed = True
            break
    # Re-merge adyacentes del mismo kind
    if not items:
        return []
    merged = [items[0]]
    for z in items[1:]:
        prev = merged[-1]
        if prev[u"kind"] == z[u"kind"] and abs(float(prev[u"u1"]) - float(z[u"u0"])) < 1e-6:
            prev[u"u1"] = float(z[u"u1"])
        else:
            merged.append(z)
    return merged


def zonas_cabeza_muro(host, walls_sel, muro_apilado_sobre_fn):
    """
    Paños a lo largo del host: ``[{u0,u1,kind}, ...]`` con kind
    ``solape`` | ``desacople``.

    Si no hay muro encima → un solo ``desacople`` [0,1].
    Si el encima cubre todo → un solo ``solape`` [0,1].
    """
    if host is None:
        return []
    covered = _intervalos_solape_u(host, walls_sel, muro_apilado_sobre_fn)
    min_u = _min_pano_u(host)
    zones = []
    for a, b in covered:
        zones.append({u"u0": float(a), u"u1": float(b), u"kind": u"solape"})
    for a, b in _complement_intervals(covered, 0.0, 1.0):
        zones.append({u"u0": float(a), u"u1": float(b), u"kind": u"desacople"})
    zones.sort(key=lambda z: (float(z[u"u0"]), float(z[u"u1"])))
    return _fuse_small_zones(zones, min_u)


def debe_partir_por_desacople(zones, terminacion_modo=None):
    """
    True si hay solape y desacople y la terminación de cabeza requiere
    comportamientos distintos (auto / empotramiento).
    """
    if not zones or len(zones) < 2:
        return False
    kinds = set(z.get(u"kind") for z in zones)
    if u"solape" not in kinds or u"desacople" not in kinds:
        return False
    modo = terminacion_modo
    if modo in (u"pata_l", u"recto"):
        # Misma terminación en todo el muro → no hace falta partir por cabeza.
        return False
    # auto (None) o empotramiento → partir
    return True


def terminacion_para_zona(kind, terminacion_modo=None):
    """
    Terminación de cabeza del paño.

    - ``recto`` / ``pata_l`` forzados: se respetan en todas las zonas.
    - ``empotramiento`` o auto: solape → empotramiento; desacople → pata_l.
    """
    if terminacion_modo == u"recto":
        return u"recto"
    if terminacion_modo == u"pata_l":
        return u"pata_l"
    if kind == u"solape":
        return u"empotramiento"
    return u"pata_l"


def curvas_rectangulo_tramo_muro(wall, u0, u1):
    """
    Polígono cerrado (4 líneas) del tramo [u0,u1] en alzado: LocationCurve × altura bbox.

    :returns: ``List[Curve]`` o ``None``.
    """
    lc = _location_curve(wall)
    if lc is None:
        return None
    try:
        ua = max(0.0, min(1.0, float(u0)))
        ub = max(0.0, min(1.0, float(u1)))
    except Exception:
        return None
    if ub - ua < 1e-6:
        return None
    try:
        p0 = lc.Evaluate(ua, True)
        p1 = lc.Evaluate(ub, True)
    except Exception:
        return None
    try:
        bb = wall.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return None
    try:
        z0 = float(bb.Min.Z)
        z1 = float(bb.Max.Z)
    except Exception:
        return None
    if abs(z1 - z0) < 1e-9:
        return None

    def _at(p, z):
        return XYZ(float(p.X), float(p.Y), float(z))

    try:
        a = _at(p0, z0)
        b = _at(p1, z0)
        c = _at(p1, z1)
        d = _at(p0, z1)
        curves = [
            Line.CreateBound(a, b),
            Line.CreateBound(b, c),
            Line.CreateBound(c, d),
            Line.CreateBound(d, a),
        ]
    except Exception:
        return None
    try:
        out = List[Curve]()
        for crv in curves:
            out.Add(crv)
        return out
    except Exception:
        return None


def _rebar_punto_ref_xy(rebar):
    if rebar is None:
        return None
    try:
        bb = rebar.get_BoundingBox(None)
        if bb is not None:
            return XYZ(
                0.5 * (float(bb.Min.X) + float(bb.Max.X)),
                0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
                0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
            )
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption

        crvs = rebar.GetCenterlineCurves(
            False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, 0,
        )
        if crvs is not None and int(crvs.Count) >= 1:
            c0 = crvs[0]
            return c0.Evaluate(0.5, True)
    except Exception:
        pass
    return None


def _u_en_covered(u, covered):
    if u is None:
        return False
    for a, b in covered or []:
        if float(a) - 1e-6 <= float(u) <= float(b) + 1e-6:
            return True
    return False


def _rebar_n_posiciones(rebar):
    if rebar is None:
        return 0
    best = 1
    for getter in (
        lambda: int(rebar.NumberOfBarPositions),
        lambda: int(rebar.GetNumberOfBarPositions()),
        lambda: int(rebar.Quantity),
    ):
        try:
            n = int(getter())
            if n > best:
                best = n
        except Exception:
            continue
    return best


def _bar_incluido(rebar, idx):
    """True si la posición está incluida (o no se puede consultar → asumir sí)."""
    if rebar is None:
        return False
    i = int(idx)
    try:
        return bool(rebar.IsBarIncluded(i))
    except Exception:
        pass
    try:
        return bool(rebar.DoesBarExistAtPosition(i))
    except Exception:
        return True


def _indices_incluidos(rebar):
    n = max(0, _rebar_n_posiciones(rebar))
    return [i for i in range(n) if _bar_incluido(rebar, i)]


def _get_bar_transform(rebar, bar_index):
    bi = int(bar_index)
    for getter in (
        lambda: rebar.GetBarPositionTransform(bi),
        lambda: rebar.GetMovedBarTransform(bi),
    ):
        try:
            t = getter()
            if t is not None:
                return t
        except Exception:
            continue
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None and hasattr(acc, "GetBarPositionTransform"):
            return acc.GetBarPositionTransform(bi)
    except Exception:
        pass
    return None


def _mid_barra_posicion(rebar, pos_idx):
    """
    Punto medio aproximado de la barra en ``pos_idx``.

    ``GetCenterlineCurves(..., pos)`` en sets de AR a menudo ignora ``pos`` y
    devuelve siempre la barra 0 — por eso se prioriza el transform de posición.
    """
    tr = _get_bar_transform(rebar, int(pos_idx))
    if tr is not None:
        try:
            return tr.Origin
        except Exception:
            pass
    # Curvas: solo fiables de forma segura en pos 0
    if int(pos_idx) == 0:
        try:
            from Autodesk.Revit.DB.Structure import MultiplanarOption

            crvs = rebar.GetCenterlineCurves(
                False, False, False,
                MultiplanarOption.IncludeAllMultiplanarCurves, 0,
            )
        except Exception:
            crvs = None
        if crvs is not None and int(crvs.Count) >= 1:
            try:
                return crvs[0].Evaluate(0.5, True)
            except Exception:
                try:
                    p0 = crvs[0].GetEndPoint(0)
                    p1 = crvs[0].GetEndPoint(1)
                    return XYZ(
                        0.5 * (float(p0.X) + float(p1.X)),
                        0.5 * (float(p0.Y) + float(p1.Y)),
                        0.5 * (float(p0.Z) + float(p1.Z)),
                    )
                except Exception:
                    pass
    if int(pos_idx) != 0:
        mid0 = _mid_barra_posicion(rebar, 0)
        t0 = _get_bar_transform(rebar, 0)
        ti = _get_bar_transform(rebar, pos_idx)
        if mid0 is not None and t0 is not None and ti is not None:
            try:
                return ti.OfPoint(t0.Inverse.OfPoint(mid0))
            except Exception:
                pass
        # Traslación por span del set (último recurso)
        return None
    return None


def _mapa_u_por_indice(rebar, host):
    """
    ``{idx: u}`` a lo largo del host.

    Si los mids/transforms no distinguen posiciones (todas el mismo u), interpola
    por el span del set: idx 0 → extremo del span según la barra 0 medida.
    """
    n = max(1, _rebar_n_posiciones(rebar))
    lc = _location_curve(host)
    tang = _tangent_xy(lc)
    measured = {}
    for i in range(n):
        mid = _mid_barra_posicion(rebar, i)
        u = _u_along_host(lc, mid, tang) if mid is not None else None
        if u is not None:
            measured[i] = float(u)
    span = _rebar_u_span_en_host(rebar, host)
    degenerate = False
    if len(measured) >= 2:
        vals = list(measured.values())
        if max(vals) - min(vals) < 1e-3:
            degenerate = True
    elif n > 1:
        degenerate = True

    if not degenerate and len(measured) >= n:
        return measured

    # Interpolar sobre el span del set
    if span is None:
        return measured
    s0, s1 = float(span[0]), float(span[1])
    u_at_0 = measured.get(0)
    u_at_last = measured.get(n - 1) if n > 1 else None
    if u_at_0 is not None and not degenerate:
        # Sentido: idx0 cerca de s0 o s1
        if abs(u_at_0 - s0) <= abs(u_at_0 - s1):
            u_end0, u_end1 = s0, s1
        else:
            u_end0, u_end1 = s1, s0
    elif u_at_0 is not None and u_at_last is not None and abs(u_at_0 - u_at_last) > 1e-3:
        u_end0, u_end1 = u_at_0, u_at_last
    else:
        # Degenerado: asumir idx crece con u (s0→s1). La barra 0 del AR
        # suele estar en un extremo; si u_at_0 ≈ s1, invertir.
        if u_at_0 is not None and abs(u_at_0 - s1) < abs(u_at_0 - s0):
            u_end0, u_end1 = s1, s0
        else:
            u_end0, u_end1 = s0, s1

    out = {}
    for i in range(n):
        if n == 1:
            out[i] = u_end0
        else:
            t = float(i) / float(n - 1)
            out[i] = u_end0 + t * (u_end1 - u_end0)
    return out


def _clasificar_posiciones_vertical(rebar, host, covered):
    """
    :returns: ``(idxs_solape, idxs_desacople)`` listas de índices de posición
    **incluidos** (ignora ``SetBarIncluded(False)``).
    """
    us_by_i = _mapa_u_por_indice(rebar, host)
    n = max(1, _rebar_n_posiciones(rebar))
    solape = []
    desacople = []
    for i in range(n):
        if not _bar_incluido(rebar, i):
            continue
        u = us_by_i.get(i)
        if u is None:
            # Sin u: no llamar a _set_cruza (recursión). Conservador → desacople.
            desacople.append(i)
            continue
        if _u_en_covered(u, covered):
            solape.append(i)
        else:
            desacople.append(i)
    return solape, desacople


def _us_de_indices(rebar, host, indices):
    """Lista de u para ``indices`` (mapa span-aware; no confiar solo en curvas)."""
    mapa = _mapa_u_por_indice(rebar, host)
    out = []
    for i in indices or []:
        u = mapa.get(int(i))
        if u is not None:
            out.append(float(u))
    return out


def _curves_barra_posicion(rebar, pos_idx, host=None):
    """Curvas de la barra en ``pos_idx`` (lista Python), o ``None``."""
    from Autodesk.Revit.DB.Structure import MultiplanarOption

    i = int(pos_idx)
    # Siempre partir de pos 0: GetCenterlineCurves(i) suele ignorar i
    try:
        crvs0 = rebar.GetCenterlineCurves(
            False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, 0,
        )
    except Exception:
        crvs0 = None
    if crvs0 is None or int(crvs0.Count) < 1:
        return None
    chain0 = [crvs0[k] for k in range(crvs0.Count)]
    if i == 0:
        return chain0

    t0 = _get_bar_transform(rebar, 0)
    ti = _get_bar_transform(rebar, i)
    if t0 is not None and ti is not None:
        try:
            # ¿Transforms distintos?
            o0 = t0.Origin
            oi = ti.Origin
            dist = float(o0.DistanceTo(oi)) if o0 is not None and oi is not None else 0.0
            if dist > 1e-6:
                rel = ti.Multiply(t0.Inverse)
                return [c.CreateTransformed(rel) for c in chain0]
        except Exception:
            pass

    # Deslizar por mapa u (span) si hay host
    if host is not None:
        mapa = _mapa_u_por_indice(rebar, host)
        u0 = mapa.get(0)
        ui = mapa.get(i)
        if u0 is not None and ui is not None and abs(ui - u0) > 1e-6:
            moved = _deslizar_cadena_a_u(chain0, host, ui)
            if moved:
                return moved
    return None


def _rebar_u_span_en_host(rebar, host):
    """``(u_min, u_max)`` del set proyectado en el eje del host, o ``None``.

    1) Barras **incluidas** (tras ``SetBarIncluded``).
    2) Si ese span es degenerado (p. ej. transforms AR dan todos u=0), bbox.
    """
    lc = _location_curve(host)
    if lc is None or rebar is None:
        return None
    tang = _tangent_xy(lc)
    us_incl = []
    n = _rebar_n_posiciones(rebar)
    for i in range(n):
        if not _bar_incluido(rebar, i):
            continue
        mid = _mid_barra_posicion(rebar, i)
        u = _u_along_host(lc, mid, tang) if mid is not None else None
        if u is not None:
            us_incl.append(float(u))
    if len(us_incl) >= 2 and (max(us_incl) - min(us_incl)) > 1e-3:
        return (min(us_incl), max(us_incl))

    us_bb = []
    try:
        bb = rebar.get_BoundingBox(None)
        if bb is not None:
            for pt in (
                XYZ(float(bb.Min.X), float(bb.Min.Y), float(bb.Min.Z)),
                XYZ(float(bb.Max.X), float(bb.Max.Y), float(bb.Max.Z)),
                XYZ(float(bb.Min.X), float(bb.Max.Y), float(bb.Min.Z)),
                XYZ(float(bb.Max.X), float(bb.Min.Y), float(bb.Max.Z)),
            ):
                u = _u_along_host(lc, pt, tang)
                if u is not None:
                    us_bb.append(float(u))
    except Exception:
        pass
    if us_bb and (max(us_bb) - min(us_bb)) > 1e-3:
        return (min(us_bb), max(us_bb))
    if us_incl:
        return (min(us_incl), max(us_incl))
    if us_bb:
        return (min(us_bb), max(us_bb))
    return None


def _set_cruza_solape_y_desacople(rebar, host, covered):
    """True si el set tiene barras incluidas en solape y en desacople, o el span cruza."""
    if not covered or rebar is None or host is None:
        return False
    # Por u de barras incluidas (sin recursión)
    us_by_i = _mapa_u_por_indice(rebar, host)
    n = max(1, _rebar_n_posiciones(rebar))
    has_sol = False
    has_des = False
    n_incl = 0
    for i in range(n):
        if not _bar_incluido(rebar, i):
            continue
        n_incl += 1
        u = us_by_i.get(i)
        if u is None:
            continue
        if _u_en_covered(u, covered):
            has_sol = True
        else:
            has_des = True
        if has_sol and has_des:
            return True
    # Tras SetBarIncluded: no usar bbox lleno si las incluidas ya son solo una zona
    if n_incl > 0 and has_sol and not has_des:
        return False
    if n_incl > 0 and has_des and not has_sol:
        return False
    span = _rebar_u_span_en_host(rebar, host)
    if span is None:
        return False
    u_a, u_b = float(span[0]), float(span[1])
    if abs(u_b - u_a) < 1e-3:
        return False
    if u_b < u_a:
        u_a, u_b = u_b, u_a
    has_solape = False
    has_des2 = False
    samples = [u_a, 0.5 * (u_a + u_b), u_b]
    if u_b - u_a > 1e-6:
        for k in range(1, 8):
            samples.append(u_a + (u_b - u_a) * (float(k) / 8.0))
    for u in samples:
        if _u_en_covered(u, covered):
            has_solape = True
        else:
            has_des2 = True
        if has_solape and has_des2:
            return True
    return False


def _mid_xyz_cadena(chain):
    if not chain:
        return None
    try:
        c0 = chain[0]
        return c0.Evaluate(0.5, True)
    except Exception:
        try:
            p0 = chain[0].GetEndPoint(0)
            p1 = chain[0].GetEndPoint(1)
            return XYZ(
                0.5 * (float(p0.X) + float(p1.X)),
                0.5 * (float(p0.Y) + float(p1.Y)),
                0.5 * (float(p0.Z) + float(p1.Z)),
            )
        except Exception:
            return None


def _deslizar_cadena_a_u(chain, host, u_target):
    """Trasladar curvas a lo largo del eje del host hasta ``u_target``."""
    if not chain or host is None or u_target is None:
        return None
    lc = _location_curve(host)
    tang = _tangent_xy(lc)
    if lc is None or tang is None:
        return None
    mid = _mid_xyz_cadena(chain)
    if mid is None:
        return None
    u_cur = _u_along_host(lc, mid, tang)
    if u_cur is None:
        return None
    try:
        L = float(lc.Length)
    except Exception:
        return None
    du = float(u_target) - float(u_cur)
    if abs(du) < 1e-12:
        return list(chain)
    try:
        delta = XYZ(
            float(tang.X) * du * L,
            float(tang.Y) * du * L,
            0.0,
        )
        tr = Transform.CreateTranslation(delta)
        return [c.CreateTransformed(tr) for c in chain]
    except Exception:
        return None


def _doc_regen(doc):
    if doc is None:
        return
    try:
        doc.Regenerate()
    except Exception:
        pass


def _spacing_rebar(rebar):
    try:
        sp = float(rebar.MaxSpacing)
        if sp > 1e-12:
            return sp
    except Exception:
        pass
    return None


def _aplicar_layout_side(acc, n, alen, sp, b_side):
    """Aplica layout; retorna ``(ok, err_str)``."""
    errs = []
    if sp is not None and sp > 1e-12:
        try:
            acc.SetLayoutAsNumberWithSpacing(
                int(n), float(sp), float(alen), bool(b_side), True, True,
            )
            return True, u""
        except Exception as ex:
            errs.append(u"NWS:{0}".format(ex))
    try:
        acc.SetLayoutAsFixedNumber(
            int(n), float(alen), bool(b_side), True, True,
        )
        return True, u""
    except Exception as ex:
        errs.append(u"FN:{0}".format(ex))
    # Sin incluir extremos a veces pasa
    try:
        acc.SetLayoutAsFixedNumber(
            int(n), float(alen), bool(b_side), True, False,
        )
        return True, u""
    except Exception as ex:
        errs.append(u"FN2:{0}".format(ex))
    return False, u" | ".join(errs) if errs else u"layout fail"


def _qty_rebar(rebar):
    for getter in (
        lambda: int(rebar.Quantity),
        lambda: int(rebar.NumberOfBarPositions),
        lambda: int(rebar.GetNumberOfBarPositions()),
    ):
        try:
            n = int(getter())
            if n > 0:
                return n
        except Exception:
            continue
    return 0


def _alen_accessor(acc):
    if acc is None:
        return 0.0
    try:
        return float(acc.ArrayLength)
    except Exception:
        try:
            return float(acc.GetArrayLength())
        except Exception:
            return 0.0


def _offset_xy_barra_vs_eje(rebar, host):
    """Offset XY de la barra 0 respecto al LocationCurve del host."""
    lc = _location_curve(host)
    tang = _tangent_xy(lc)
    if lc is None or tang is None:
        return XYZ(0.0, 0.0, 0.0)
    mid = _mid_barra_posicion(rebar, 0)
    if mid is None:
        chain = _curves_barra_posicion(rebar, 0, host=host)
        mid = _mid_xyz_cadena(chain) if chain else None
    if mid is None:
        return XYZ(0.0, 0.0, 0.0)
    u = _u_along_host(lc, mid, tang)
    if u is None:
        return XYZ(0.0, 0.0, 0.0)
    try:
        axis_pt = lc.Evaluate(float(u), True)
        return XYZ(
            float(mid.X) - float(axis_pt.X),
            float(mid.Y) - float(axis_pt.Y),
            0.0,
        )
    except Exception:
        return XYZ(0.0, 0.0, 0.0)


def _bar_z_extent(rebar):
    """``(z_lo, z_hi)`` de la barra 0."""
    chain = _curves_barra_posicion(rebar, 0)
    if not chain:
        try:
            from Autodesk.Revit.DB.Structure import MultiplanarOption

            crvs = rebar.GetCenterlineCurves(
                False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, 0,
            )
            chain = [crvs[i] for i in range(crvs.Count)] if crvs else None
        except Exception:
            chain = None
    zs = []
    for c in chain or []:
        try:
            zs.append(float(c.GetEndPoint(0).Z))
            zs.append(float(c.GetEndPoint(1).Z))
        except Exception:
            pass
    if not zs:
        return None, None
    return min(zs), max(zs)


def _scale_copy_to_u_zone(doc, new_rb, rebar_src, host, u_lo, u_hi, n, b_side):
    """
    1) SetLayout a ``n`` barras (reduce cantidad; el span puede seguir lleno por constraints AR).
    2) ``ScaleToBox`` en barra 0 con ancho de zona (ahora sí con n barras).
    3) ``Move`` hasta ``u_lo``.

    ScaleToBox con las 31 barras originales y SetLayout(16) después colapsaba el solape.
    """
    lc = _location_curve(host)
    tang = _tangent_xy(lc)
    if lc is None or tang is None:
        return False, u"sin eje"
    try:
        L = float(lc.Length)
    except Exception:
        return False, u"Length"
    try:
        acc = new_rb.GetShapeDrivenAccessor()
    except Exception as ex:
        return False, u"accessor: {0}".format(ex)
    if acc is None:
        return False, u"accessor null"

    offset = _offset_xy_barra_vs_eje(rebar_src, host)
    z_lo, z_hi = _bar_z_extent(rebar_src)
    if z_lo is None or z_hi is None or abs(z_hi - z_lo) < 1e-9:
        return False, u"sin extensión Z barra"

    mapa = _mapa_u_por_indice(rebar_src, host)
    u_bar0 = float(mapa.get(0, 0.0))
    try:
        p_bar0 = lc.Evaluate(float(u_bar0), True)
    except Exception as ex:
        return False, u"Evaluate u_bar0: {0}".format(ex)

    origin = XYZ(
        float(p_bar0.X) + float(offset.X),
        float(p_bar0.Y) + float(offset.Y),
        float(z_lo),
    )
    width_u = abs(float(u_hi) - float(u_lo))
    sp = _spacing_rebar(rebar_src)
    # Longitud de reparto: preferir (n-1)*spacing del original (coherente con AR)
    if sp is not None and n > 1:
        dist = float(sp) * float(n - 1)
    else:
        dist = width_u * L
    if dist < 1e-9:
        return False, u"dist=0"

    x_dist = XYZ(float(tang.X) * dist, float(tang.Y) * dist, 0.0)
    y_bar = XYZ(0.0, 0.0, float(z_hi) - float(z_lo))
    u_mid = 0.5 * (float(u_lo) + float(u_hi))
    # ArrayLength “lleno” del original para el primer SetLayout
    try:
        alen_full = _alen_accessor(rebar_src.GetShapeDrivenAccessor())
    except Exception:
        alen_full = L
    if alen_full < 1e-9:
        alen_full = L

    last = u""
    for b_try in (b_side, (not b_side)):
        # --- A) Reducir cantidad primero (span puede seguir [0,1]) ---
        ok_q, err_q = _aplicar_layout_side(acc, n, alen_full, sp, b_try)
        if not ok_q and sp is not None and n > 1:
            ok_q, err_q = _aplicar_layout_side(
                acc, n, float(sp) * float(n - 1), sp, b_try,
            )
        if not ok_q:
            last = u"SetLayout qty: {0}".format(err_q)
            continue
        _doc_regen(doc)
        qty1 = _qty_rebar(new_rb)
        if qty1 < max(1, n - 2):
            last = u"qty tras layout={0} want={1}".format(qty1, n)
            continue

        # --- B) ScaleToBox al ancho de zona, anclado en barra 0 ---
        try:
            acc.ScaleToBox(origin, x_dist, y_bar)
        except Exception as ex_sc:
            last = u"ScaleToBox: {0}".format(ex_sc)
            continue
        _doc_regen(doc)

        # Reafirmar layout con alen = dist de zona
        ok_lay, err_lay = _aplicar_layout_side(acc, n, dist, sp, b_try)
        if not ok_lay:
            last = u"layout post-scale: {0}".format(err_lay)
            continue
        _doc_regen(doc)

        span1 = _rebar_u_span_en_host(new_rb, host)
        if span1 is None:
            last = u"sin span tras scale"
            continue
        wu1 = abs(float(span1[1]) - float(span1[0]))
        min_wu = max(0.02, 0.30 * (dist / L if L > 1e-12 else width_u))
        if wu1 < min_wu:
            last = u"span tras scale estrecho {0} qty={1}".format(
                _fmt_intervals([span1]), _qty_rebar(new_rb),
            )
            continue
        if wu1 > 0.90 and width_u < 0.85:
            last = u"span tras scale sigue lleno {0}".format(_fmt_intervals([span1]))
            continue

        # --- C) Mover a u_lo ---
        span_min = min(float(span1[0]), float(span1[1]))
        du = float(u_lo) - span_min
        if abs(du) > 1e-8:
            try:
                from Autodesk.Revit.DB import ElementTransformUtils

                ElementTransformUtils.MoveElement(
                    doc,
                    new_rb.Id,
                    XYZ(float(tang.X) * du * L, float(tang.Y) * du * L, 0.0),
                )
            except Exception as ex_mv:
                last = u"move a u_lo: {0}".format(ex_mv)
                continue
            _doc_regen(doc)

        span2 = _rebar_u_span_en_host(new_rb, host)
        if span2 is None:
            last = u"sin span tras move"
            continue
        mid2 = 0.5 * (float(span2[0]) + float(span2[1]))
        wu2 = abs(float(span2[1]) - float(span2[0]))
        qty = _qty_rebar(new_rb)
        if wu2 < min_wu:
            last = u"span tras move estrecho {0}".format(_fmt_intervals([span2]))
            continue
        if abs(mid2 - u_mid) > 0.22:
            du2 = u_mid - mid2
            try:
                from Autodesk.Revit.DB import ElementTransformUtils

                ElementTransformUtils.MoveElement(
                    doc,
                    new_rb.Id,
                    XYZ(float(tang.X) * du2 * L, float(tang.Y) * du2 * L, 0.0),
                )
                _doc_regen(doc)
                span2 = _rebar_u_span_en_host(new_rb, host)
                if span2 is not None:
                    mid2 = 0.5 * (float(span2[0]) + float(span2[1]))
            except Exception:
                pass
        if abs(mid2 - u_mid) > 0.22:
            last = u"mid={0:.3f} want={1:.3f} span={2} qty={3}".format(
                mid2, u_mid, _fmt_intervals([span2]), qty,
            )
            continue
        return True, u"qty={0} span={1} side={2}".format(
            qty, _fmt_intervals([span2]), b_try,
        )
    return False, last or u"ScaleToBox+Move fail"


def _crear_subset_via_curves(
    doc, rebar, host, u_lo, u_hi, n, create_fn, hook_orient_fn, normal_fn, b_side,
):
    """
    CreateFromCurves en la barra 0 (u_bar0) con ancho de zona, luego Move a u_lo.
    Misma estrategia que ScaleToBox+Move (el anclaje directo en u_lo falla).
    """
    if create_fn is None:
        return None, u"sin create_fn"
    lc = _location_curve(host)
    tang = _tangent_xy(lc)
    if lc is None or tang is None:
        return None, u"sin eje"
    try:
        L = float(lc.Length)
    except Exception:
        return None, u"Length"

    mapa = _mapa_u_por_indice(rebar, host)
    u_bar0 = float(mapa.get(0, 0.0))
    chain = _curves_barra_posicion(rebar, 0, host=host)
    if not chain:
        return None, u"sin curvas"
    # Quedarse en u_bar0 (no deslizar a u_lo aquí)
    chain = _deslizar_cadena_a_u(chain, host, u_bar0)
    if not chain:
        return None, u"deslizar u_bar0 fail"

    from Autodesk.Revit.DB.Structure import RebarBarType, RebarStyle

    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return None, u"bar type"
    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    try:
        o0 = hook_orient_fn(rebar, 0) if hook_orient_fn else None
        o1 = hook_orient_fn(rebar, 1) if hook_orient_fn else None
    except Exception:
        o0 = o1 = None

    norms = []
    try:
        wo = host.Orientation
        if wo is not None and float(wo.GetLength()) > 1e-12:
            wn = wo.Normalize()
            norms.append(wn)
            norms.append(XYZ(-float(wn.X), -float(wn.Y), -float(wn.Z)))
    except Exception:
        pass
    try:
        rn = normal_fn(rebar) if normal_fn else rebar.GetShapeDrivenAccessor().Normal
        if rn is not None and float(rn.GetLength()) > 1e-12:
            norms.append(rn.Normalize())
    except Exception:
        pass
    if not norms:
        return None, u"sin normales"

    width_u = abs(float(u_hi) - float(u_lo))
    alen = width_u * L
    sp = _spacing_rebar(rebar)
    u_mid = 0.5 * (float(u_lo) + float(u_hi))
    last = u""
    seen = set()
    for norm in norms:
        key = (round(float(norm.X), 5), round(float(norm.Y), 5), round(float(norm.Z), 5))
        if key in seen:
            continue
        seen.add(key)
        try:
            new_rb = create_fn(doc, chain, host, norm, bar_type, style, o0, o1)
        except Exception as ex:
            last = u"CreateFromCurves: {0}".format(ex)
            continue
        if new_rb is None:
            last = u"CreateFromCurves None"
            continue
        if n <= 1:
            du = float(u_lo) - u_bar0
            if abs(du) > 1e-8:
                try:
                    from Autodesk.Revit.DB import ElementTransformUtils

                    ElementTransformUtils.MoveElement(
                        doc,
                        new_rb.Id,
                        XYZ(float(tang.X) * du * L, float(tang.Y) * du * L, 0.0),
                    )
                except Exception as ex_m:
                    try:
                        doc.Delete(new_rb.Id)
                    except Exception:
                        pass
                    last = unicode(ex_m)
                    continue
            _doc_regen(doc)
            return new_rb, u"single"
        try:
            acc = new_rb.GetShapeDrivenAccessor()
        except Exception:
            acc = None
        if acc is None:
            try:
                doc.Delete(new_rb.Id)
            except Exception:
                pass
            last = u"no accessor"
            continue
        ok_lay = False
        for b_try in (b_side, (not b_side)):
            ok_lay, err_lay = _aplicar_layout_side(acc, n, alen, sp, b_try)
            if ok_lay:
                break
        if not ok_lay:
            try:
                doc.Delete(new_rb.Id)
            except Exception:
                pass
            last = err_lay
            continue
        _doc_regen(doc)
        span1 = _rebar_u_span_en_host(new_rb, host)
        if span1 is None or abs(float(span1[1]) - float(span1[0])) < 0.02:
            try:
                doc.Delete(new_rb.Id)
            except Exception:
                pass
            last = u"span colapsado tras layout"
            continue
        span_min = min(float(span1[0]), float(span1[1]))
        du = float(u_lo) - span_min
        if abs(du) > 1e-8:
            try:
                from Autodesk.Revit.DB import ElementTransformUtils

                ElementTransformUtils.MoveElement(
                    doc,
                    new_rb.Id,
                    XYZ(float(tang.X) * du * L, float(tang.Y) * du * L, 0.0),
                )
            except Exception as ex_mv:
                try:
                    doc.Delete(new_rb.Id)
                except Exception:
                    pass
                last = u"move: {0}".format(ex_mv)
                continue
            _doc_regen(doc)
        span = _rebar_u_span_en_host(new_rb, host)
        qty = _qty_rebar(new_rb)
        if span is None:
            try:
                doc.Delete(new_rb.Id)
            except Exception:
                pass
            last = u"sin span final"
            continue
        mid = 0.5 * (float(span[0]) + float(span[1]))
        wu = abs(float(span[1]) - float(span[0]))
        if wu < max(0.02, 0.35 * width_u) or abs(mid - u_mid) > 0.22:
            try:
                doc.Delete(new_rb.Id)
            except Exception:
                pass
            last = u"mid={0:.3f} want={1:.3f} span={2}".format(
                mid, u_mid, _fmt_intervals([span]),
            )
            continue
        return new_rb, u"curves qty={0} span={1}".format(qty, _fmt_intervals([span]))
    return None, last or u"curves fail"


def rebar_en_solape_apilado(rebar, host, walls_sel, muro_apilado_sobre_fn):
    """
    True si el rebar/set debe empotrarse (solo solape).

    - Sin muro encima / solo desacople → False (pata L).
    - Solo solape → True.
    - Set **mixto** (aún sin partir) → False (no empujar barras del desacople).
    """
    if rebar is None or host is None:
        return False
    covered = _intervalos_solape_u(host, walls_sel, muro_apilado_sobre_fn)
    if not covered:
        return False
    # Barras incluidas primero (válido tras SetBarIncluded en el subset solape)
    solape_idx, des_idx = _clasificar_posiciones_vertical(rebar, host, covered)
    if solape_idx and des_idx:
        return False
    if solape_idx and not des_idx:
        return True
    if des_idx and not solape_idx:
        return False
    if _set_cruza_solape_y_desacople(rebar, host, covered):
        return False
    pt = _rebar_punto_ref_xy(rebar)
    if pt is None:
        return False
    lc = _location_curve(host)
    u = _u_along_host(lc, pt)
    return _u_en_covered(u, covered)


def _stamp_vertical_safe(rebar):
    try:
        from armado_muros_rebar_params import stamp_malla_vertical_rebar

        stamp_malla_vertical_rebar(rebar)
    except Exception:
        pass


def _excluir_indices_en_rebar(doc, rebar, indices):
    """Excluye posiciones con SetBarIncluded; True si al menos una quedo fuera."""
    if rebar is None or not indices:
        return False
    try:
        from armado_muros_rebar_layout import _excluir_barras_por_indices

        return bool(
            _excluir_barras_por_indices(
                rebar, indices, doc=doc, regenerate=True,
            )
        )
    except Exception:
        pass
    ok = False
    for idx in sorted(set(int(i) for i in indices), reverse=True):
        try:
            rebar.SetBarIncluded(False, int(idx))
            ok = True
        except Exception:
            pass
    if ok and doc is not None:
        _doc_regen(doc)
    return ok


def _solape_desde_clon_desacople(
    doc, rb_des, rebar_src, host, u_lo, u_hi, n, covered=None,
    des_idx=None, solape_idx=None,
):
    """
    Solape por SetBarIncluded: copia el AR original y excluye barras de desacople.

    ScaleToBox solo es estable anclado en barra 0 (desacople). Anclar en u_lo/u_hi
    o Single+SetLayout colapsa / llena el host. Las barras de solape ya estan en
    su sitio; basta con apagar las del desacople (mismo patron que cabezal).
    """
    if doc is None or host is None:
        return None, u"args"
    src = rebar_src if rebar_src is not None else rb_des
    if src is None:
        return None, u"sin fuente"

    from Autodesk.Revit.DB import ElementTransformUtils
    from Autodesk.Revit.DB.Structure import Rebar

    to_exclude = list(des_idx) if des_idx else []
    if not to_exclude and solape_idx:
        keep = set(int(i) for i in solape_idx)
        npos = _rebar_n_posiciones(src)
        to_exclude = [i for i in range(npos) if i not in keep]
    if not to_exclude:
        return None, u"sin indices desacople a excluir"

    try:
        copied = ElementTransformUtils.CopyElement(
            doc, src.Id, XYZ(0.0, 0.0, 0.0),
        )
    except Exception as ex:
        return None, u"Copy: {0}".format(ex)
    if copied is None or int(copied.Count) < 1:
        return None, u"Copy vacio"
    try:
        new_rb = doc.GetElement(copied[0])
    except Exception:
        new_rb = None
    if new_rb is None or not isinstance(new_rb, Rebar):
        return None, u"copia no Rebar"

    if not _excluir_indices_en_rebar(doc, new_rb, to_exclude):
        try:
            doc.Delete(new_rb.Id)
        except Exception:
            pass
        return None, u"SetBarIncluded no excluyo"

    _doc_regen(doc)
    incl = _indices_incluidos(new_rb)
    if not incl:
        try:
            doc.Delete(new_rb.Id)
        except Exception:
            pass
        return None, u"ninguna barra incluida tras exclude"

    mapa = _mapa_u_por_indice(new_rb, host)
    us_ok = []
    us_bad = []
    for i in incl:
        u = mapa.get(int(i))
        if u is None:
            continue
        if covered is not None:
            if _u_en_covered(u, covered):
                us_ok.append(float(u))
            else:
                us_bad.append(float(u))
        else:
            us_ok.append(float(u))

    if covered is not None and len(us_ok) < max(1, int(0.6 * len(incl))):
        try:
            doc.Delete(new_rb.Id)
        except Exception:
            pass
        return None, u"incluidas fuera de solape ok={0} bad={1} n_incl={2}".format(
            len(us_ok), len(us_bad), len(incl),
        )

    span = _rebar_u_span_en_host(new_rb, host)
    _stamp_vertical_safe(new_rb)
    return new_rb, u"exclude-des n_excl={0} n_incl={1} span={2}".format(
        len(to_exclude),
        len(incl),
        _fmt_intervals([span]) if span else u"—",
    )


def _crear_subset_vertical_desde_rebar(
    doc,
    rebar,
    host,
    indices,
    create_fn,
    copy_layout_fn,
    hook_orient_fn,
    normal_fn,
    covered=None,
    expect_solape=None,
):
    """
    Subset en [u_lo,u_hi]:
    1) Copy + ScaleToBox (fiable con constraints de AR)
    2) Fallback CreateFromCurves + normal de muro
    """
    global _LAST_CREATE_FAIL
    _LAST_CREATE_FAIL = u""
    zona = u"solape" if expect_solape else (u"desacople" if expect_solape is False else u"?")

    def _fail(reason):
        global _LAST_CREATE_FAIL
        _LAST_CREATE_FAIL = u"{0}: {1}".format(zona, reason)
        return None

    if not indices or doc is None or rebar is None or host is None:
        return _fail(u"args nulos o sin índices")

    from Autodesk.Revit.DB import ElementTransformUtils
    from Autodesk.Revit.DB.Structure import Rebar

    idxs = sorted(int(i) for i in indices)
    us = _us_de_indices(rebar, host, idxs)
    if not us:
        return _fail(u"sin u de posiciones {0}".format(_fmt_idx_preview(idxs)))
    u_lo = min(us)
    u_hi = max(us)
    u_mid = 0.5 * (u_lo + u_hi)
    n = len(idxs)
    if n > 1 and abs(u_hi - u_lo) < 1e-4:
        return _fail(u"u degenerado u_lo=u_hi={0:.3f}".format(u_lo))

    sides = []
    try:
        b_src = bool(rebar.GetShapeDrivenAccessor().BarsOnNormalSide)
        sides = [b_src, (not b_src)]
    except Exception:
        sides = [True, False]

    fails = []

    # --- 1) Copy + ScaleToBox ---
    for b_side in sides:
        try:
            copied = ElementTransformUtils.CopyElement(
                doc, rebar.Id, XYZ(0.0, 0.0, 0.0),
            )
        except Exception as ex_cp:
            fails.append(u"Copy: {0}".format(ex_cp))
            break
        if copied is None or int(copied.Count) < 1:
            fails.append(u"Copy vacío")
            break
        try:
            new_rb = doc.GetElement(copied[0])
        except Exception:
            new_rb = None
        if new_rb is None or not isinstance(new_rb, Rebar):
            fails.append(u"copia no Rebar")
            break

        if n == 1:
            try:
                new_rb.GetShapeDrivenAccessor().SetLayoutAsSingle()
            except Exception:
                pass
            ok_sc, det = _scale_copy_to_u_zone(
                doc, new_rb, rebar, host, u_lo, u_hi, 1, b_side,
            )
            # For n=1 ScaleToBox still places the bar; if fails, move manually
            if not ok_sc:
                mapa = _mapa_u_por_indice(rebar, host)
                u0 = mapa.get(0, 0.0)
                lc = _location_curve(host)
                tang = _tangent_xy(lc)
                L = float(lc.Length)
                du = float(u_lo) - float(u0)
                if tang is not None and abs(du) > 1e-9:
                    try:
                        ElementTransformUtils.MoveElement(
                            doc,
                            new_rb.Id,
                            XYZ(float(tang.X) * du * L, float(tang.Y) * du * L, 0.0),
                        )
                        _doc_regen(doc)
                        ok_sc, det = True, u"single move"
                    except Exception as ex_m:
                        ok_sc, det = False, unicode(ex_m)
        else:
            ok_sc, det = _scale_copy_to_u_zone(
                doc, new_rb, rebar, host, u_lo, u_hi, n, b_side,
            )

        if not ok_sc:
            fails.append(u"Scale side={0}: {1}".format(b_side, det))
            try:
                doc.Delete(new_rb.Id)
            except Exception:
                pass
            continue

        if covered is not None and expect_solape is not None:
            span = _rebar_u_span_en_host(new_rb, host)
            if span is None:
                fails.append(u"Scale side={0}: sin span final".format(b_side))
                try:
                    doc.Delete(new_rb.Id)
                except Exception:
                    pass
                continue
            mid = 0.5 * (float(span[0]) + float(span[1]))
            in_sol = _u_en_covered(mid, covered)
            wu = abs(float(span[1]) - float(span[0]))
            if n > 1 and wu < 0.02:
                fails.append(u"Scale side={0}: colapsado".format(b_side))
                try:
                    doc.Delete(new_rb.Id)
                except Exception:
                    pass
                continue
            if bool(expect_solape) != bool(in_sol):
                fails.append(
                    u"Scale side={0}: zona mid={1:.3f} span={2} in_sol={3}".format(
                        b_side, mid, _fmt_intervals([span]), in_sol,
                    ),
                )
                try:
                    doc.Delete(new_rb.Id)
                except Exception:
                    pass
                continue

        try:
            from armado_muros_rebar_params import stamp_malla_vertical_rebar

            stamp_malla_vertical_rebar(new_rb)
        except Exception:
            pass
        _LAST_CREATE_FAIL = u""
        return new_rb

    # --- 2) Fallback CreateFromCurves ---
    for b_side in sides:
        new_rb, det = _crear_subset_via_curves(
            doc, rebar, host, u_lo, u_hi, n,
            create_fn, hook_orient_fn, normal_fn, b_side,
        )
        if new_rb is None:
            fails.append(u"Curves side={0}: {1}".format(b_side, det))
            continue
        if covered is not None and expect_solape is not None:
            span = _rebar_u_span_en_host(new_rb, host)
            mid = 0.5 * (float(span[0]) + float(span[1])) if span else None
            in_sol = _u_en_covered(mid, covered) if mid is not None else None
            if mid is None or bool(expect_solape) != bool(in_sol):
                fails.append(
                    u"Curves side={0}: zona mid={1} in_sol={2}".format(
                        b_side, mid, in_sol,
                    ),
                )
                try:
                    doc.Delete(new_rb.Id)
                except Exception:
                    pass
                continue
        try:
            from armado_muros_rebar_params import stamp_malla_vertical_rebar

            stamp_malla_vertical_rebar(new_rb)
        except Exception:
            pass
        _LAST_CREATE_FAIL = u""
        return new_rb

    return _fail(u" ; ".join(fails) if fails else u"sin método válido")


def _indice_corte_desacople_prefijo(des_idx, solape_idx, npos):
    """
    Índice de corte para DividirRebarSet: último índice del prefijo desacople.

    Requiere desacople = ``0..k`` contiguo y solape = ``k+1..`` (caso apilado típico).
    :returns: ``(ok, cut_idx_or_err)``
    """
    if not des_idx or not solape_idx:
        return False, u"sin índices"
    des_s = sorted(int(i) for i in des_idx)
    sol_s = sorted(int(i) for i in solape_idx)
    # Prefijo desde 0
    for expect, got in enumerate(des_s):
        if got != expect:
            return False, u"desacople no es prefijo 0..k: {0}".format(des_s[:8])
    cut = des_s[-1]
    if sol_s[0] != cut + 1:
        return False, u"solape no continúa tras des (cut={0} sol0={1})".format(
            cut, sol_s[0],
        )
    n = int(npos)
    if cut > n - 2:
        return False, u"cut={0} inválido para n={1}".format(cut, n)
    return True, cut


def partir_rebar_vertical_si_mixto(
    doc,
    rebar,
    host,
    walls_sel,
    muro_apilado_sobre_fn,
    create_fn=None,
    copy_layout_fn=None,
    hook_orient_fn=None,
    normal_fn=None,
    messages=None,
):
    """
    Si el set vertical tiene barras en solape y desacople, lo divide con el
    patrón DividirRebarSet (Copy+Move+SetLayout). Si no hay mezcla, ``[rebar]``.

    :returns: ``(lista_rebars, n_partidos)`` — n_partidos 1 si se partió.
    """
    sink = messages if messages is not None else []
    rid = _element_id_int(getattr(rebar, "Id", None)) if rebar is not None else None
    hid = _element_id_int(getattr(host, "Id", None)) if host is not None else None

    if doc is None or rebar is None or host is None:
        log_desacople(u"skip: doc/rebar/host nulo", sink)
        return [rebar] if rebar is not None else [], 0

    covered = _intervalos_solape_u(host, walls_sel, muro_apilado_sobre_fn)
    if not covered:
        n_sobre = 0
        for other in walls_sel or []:
            try:
                if other is not None and muro_apilado_sobre_fn(host, other):
                    n_sobre += 1
            except Exception:
                pass
        log_desacople(
            u"muro {0} rebar {1}: sin solape (muros_sobre={2}) → no partir".format(
                hid, rid, n_sobre,
            ),
            sink,
        )
        return [rebar], 0

    zones = []
    for a, b in covered:
        zones.append((a, b, u"solape"))
    for a, b in _complement_intervals(covered, 0.0, 1.0):
        zones.append((a, b, u"desacople"))
    kinds = set(z[2] for z in zones)
    log_desacople(
        u"muro {0} rebar {1}: covered={2} zonas={3}".format(
            hid, rid, _fmt_intervals(covered),
            u",".join(u"{0}:{1:.2f}-{2:.2f}".format(z[2], z[0], z[1]) for z in zones),
        ),
        sink,
        to_ui=False,
    )
    if u"solape" not in kinds or u"desacople" not in kinds:
        log_desacople(
            u"muro {0} rebar {1}: solo zona {2} → no partir".format(
                hid, rid, u"/".join(sorted(kinds)),
            ),
            sink,
        )
        return [rebar], 0

    npos = _rebar_n_posiciones(rebar)
    span0 = _rebar_u_span_en_host(rebar, host)
    solape_idx, des_idx = _clasificar_posiciones_vertical(rebar, host, covered)
    log_desacople(
        u"muro {0} rebar {1}: npos={2} span={3} solape_n={4} des_n={5} sol={6} des={7}".format(
            hid, rid, npos,
            _fmt_intervals([span0]) if span0 else u"—",
            len(solape_idx or []), len(des_idx or []),
            _fmt_idx_preview(solape_idx), _fmt_idx_preview(des_idx),
        ),
        sink,
        to_ui=False,
    )
    if (not solape_idx or not des_idx) and _set_cruza_solape_y_desacople(
        rebar, host, covered,
    ):
        span = _rebar_u_span_en_host(rebar, host)
        n = max(1, npos)
        if span is not None and n > 1:
            u0, u1 = float(span[0]), float(span[1])
            solape_idx, des_idx = [], []
            for i in range(n):
                t = float(i) / float(n - 1) if n > 1 else 0.5
                u = u0 + t * (u1 - u0)
                if _u_en_covered(u, covered):
                    solape_idx.append(i)
                else:
                    des_idx.append(i)
            log_desacople(
                u"muro {0} rebar {1}: reclasificado por span → solape_n={2} des_n={3}".format(
                    hid, rid, len(solape_idx), len(des_idx),
                ),
                sink,
                to_ui=False,
            )
    if not solape_idx or not des_idx:
        log_desacople(
            u"muro {0} rebar {1}: sin mezcla solape_n={2} des_n={3} → no partir".format(
                hid, rid, len(solape_idx or []), len(des_idx or []),
            ),
            sink,
        )
        return [rebar], 0

    # Caso: desacople al final (solape prefijo) — invertir sentido del corte
    # invirtiendo índices conceptualmente: cortar tras último solape.
    ok_cut, cut_or_err = _indice_corte_desacople_prefijo(des_idx, solape_idx, npos)
    cut_idx = cut_or_err
    des_is_left = True
    if not ok_cut:
        # ¿Solape es prefijo 0..k y desacople el resto?
        ok_cut2, cut2 = _indice_corte_desacople_prefijo(solape_idx, des_idx, npos)
        if ok_cut2:
            ok_cut = True
            cut_idx = cut2
            des_is_left = False
            log_desacople(
                u"muro {0} rebar {1}: solape es prefijo → cut={2} (des a la derecha)".format(
                    hid, rid, cut_idx,
                ),
                sink,
                to_ui=False,
            )
        else:
            log_desacople(
                u"muro {0} rebar {1}: no se puede cortar — {2}".format(
                    hid, rid, cut_or_err,
                ),
                sink,
            )
            return [rebar], 0

    try:
        from dividir_rebar_set_max_spacing import dividir_rebar_set_en_indice_en_tx
    except Exception as ex_imp:
        log_desacople(
            u"muro {0} rebar {1}: import dividir_rebar_set: {2}".format(
                hid, rid, ex_imp,
            ),
            sink,
        )
        return [rebar], 0

    ok_sp, msg_sp, rbs = dividir_rebar_set_en_indice_en_tx(doc, rebar, cut_idx)
    if not ok_sp or not rbs or len(rbs) < 2:
        log_desacople(
            u"muro {0} rebar {1}: split FAIL — {2}".format(
                hid, rid, msg_sp or u"sin rebars",
            ),
            sink,
        )
        return [rebar], 0

    rb_a, rb_b = rbs[0], rbs[1]
    if des_is_left:
        rb_des, rb_sol = rb_a, rb_b
    else:
        rb_sol, rb_des = rb_a, rb_b

    for rb, tag in ((rb_des, u"desacople"), (rb_sol, u"solape")):
        try:
            from armado_muros_rebar_params import stamp_malla_vertical_rebar

            stamp_malla_vertical_rebar(rb)
        except Exception:
            pass
        sp = _rebar_u_span_en_host(rb, host)
        log_desacople(
            u"muro {0} rebar {1}: subset {2} OK id={3} span={4}".format(
                hid, rid, tag,
                _element_id_int(getattr(rb, "Id", None)),
                _fmt_intervals([sp]) if sp else u"—",
            ),
            sink,
            to_ui=False,
        )

    log_desacople(
        u"muro {0} rebar {1}: PARTIDO OK via split cut={2} → {3} ({4})".format(
            hid, rid, cut_idx, msg_sp or u"ok",
            u"des|sol" if des_is_left else u"sol|des",
        ),
        sink,
    )
    return [rb_des, rb_sol], 1


def partir_verticales_lote_por_desacople(
    doc,
    walls,
    rebars_por_muro_id,
    muro_apilado_sobre_fn,
    es_vertical_cara_fn=None,
):
    """
    Recorre ``rebars_por_muro_id`` y parte sets verticales mixtos in-place.

    :returns: ``{u"n_partidos": int, u"messages": [...]}``
    """
    res = {u"n_partidos": 0, u"messages": []}
    if not ENABLE_DESACOPLE:
        return res
    if doc is None or not rebars_por_muro_id:
        return res

    wall_by_id = {}
    for w in walls or []:
        try:
            wid = _element_id_int(getattr(w, "Id", None))
            if wid is not None:
                wall_by_id[int(wid)] = w
        except Exception:
            pass

    create_fn = copy_layout_fn = hook_orient_fn = normal_fn = None
    try:
        from arearein_verticales_empotramiento_rps import (
            _copy_layout_rebar_shape_driven,
            _create_from_curves_no_hooks,
            _hook_orient_for_create,
            _rebar_normal,
        )

        create_fn = _create_from_curves_no_hooks
        copy_layout_fn = _copy_layout_rebar_shape_driven
        hook_orient_fn = _hook_orient_for_create
        normal_fn = _rebar_normal
    except Exception as ex_h:
        res[u"messages"].append(
            u"Partir verticales desacople: helpers CreateFromCurves no disponibles.",
        )
        log_desacople(u"lote: helpers fail {0}".format(ex_h), res[u"messages"])
        return res

    from Autodesk.Revit.DB.Structure import Rebar

    n_seen = 0
    n_skip_cara = 0
    for wid, eid_list in list((rebars_por_muro_id or {}).items()):
        host = wall_by_id.get(int(wid))
        if host is None:
            log_desacople(
                u"lote: muro id {0} no en selección".format(wid),
                res[u"messages"],
            )
            continue
        # Mutar la lista in-place: ``rebars_lote`` en post-proceso suele
        # compartir la misma referencia; un ``dict[wid]=new_list`` dejaría
        # el stamp/spacing sobre IDs viejos y las etiquetas sin Orientacion.
        new_list = []
        for eid in list(eid_list or []):
            rebar = doc.GetElement(eid)
            if rebar is None or not isinstance(rebar, Rebar):
                if eid is not None:
                    new_list.append(eid)
                continue
            n_seen += 1
            if es_vertical_cara_fn is not None:
                try:
                    if not es_vertical_cara_fn(rebar, host):
                        n_skip_cara += 1
                        new_list.append(rebar.Id)
                        continue
                except Exception as ex_f:
                    log_desacople(
                        u"lote: filtro cara ex muro {0}: {1}".format(wid, ex_f),
                        res[u"messages"],
                    )
                    new_list.append(rebar.Id)
                    continue
            parts, n_p = partir_rebar_vertical_si_mixto(
                doc,
                rebar,
                host,
                walls,
                muro_apilado_sobre_fn,
                create_fn=create_fn,
                copy_layout_fn=copy_layout_fn,
                hook_orient_fn=hook_orient_fn,
                normal_fn=normal_fn,
                messages=res[u"messages"],
            )
            res[u"n_partidos"] += int(n_p)
            for rb in parts or []:
                if rb is not None:
                    try:
                        new_list.append(rb.Id)
                    except Exception:
                        pass
        try:
            eid_list[:] = new_list
        except Exception:
            rebars_por_muro_id[wid] = new_list

    log_desacople(
        u"lote fin: verticales_vistos={0} skip_cara={1} partidos={2}".format(
            n_seen, n_skip_cara, int(res[u"n_partidos"]),
        ),
        res[u"messages"],
    )
    return res
