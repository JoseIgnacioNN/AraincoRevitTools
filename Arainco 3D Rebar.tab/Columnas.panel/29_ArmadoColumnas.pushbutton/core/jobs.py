# -*- coding: utf-8 -*-
"""
Generación de rejilla de barras y fusión vertical de tramos.

REGLA DE CAPA:
- No crea ningún elemento Revit ni abre transacciones.
- No importa módulos de ui/ ni creators/.
- Recibe datos de core/geometry.py y devuelve estructuras Python puras.

Portado literalmente desde column_reinforcement_layout_rps.py.
"""
from __future__ import print_function
import math
from collections import defaultdict

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import XYZ

from core.geometry import _element_id_iv

XY_KEY_DECIMALS_DEFAULT = 9


# ---------------------------------------------------------------------------
# Posiciones lineales
# ---------------------------------------------------------------------------

def get_positions(length, count, edge_cover):
    if count < 2:
        return []
    if count == 2:
        return [-length / 2.0 + edge_cover, length / 2.0 - edge_cover]
    span = length - (2.0 * edge_cover)
    spacing = span / float(count - 1)
    return [
        -length / 2.0 + edge_cover + (i * spacing)
        for i in range(count)
    ]


# ---------------------------------------------------------------------------
# Conteo de barras de perímetro
# ---------------------------------------------------------------------------

def perimeter_outer_bar_count(bars_a, bars_b):
    if bars_a < 2 or bars_b < 2:
        return 0
    return (2 * bars_a) + (2 * bars_b) - 4


def perimeter_inner_outline_count(bars_a, bars_b):
    if bars_a < 4 or bars_b < 4:
        return 0
    ia = bars_a - 2
    ib = bars_b - 2
    if ia < 2 or ib < 2:
        return 0
    return (2 * ia) + (2 * ib) - 4


def hilos_esperados_una_columna(bars_a, bars_b, include_inner_outline):
    outer_n = perimeter_outer_bar_count(bars_a, bars_b)
    inner_n = perimeter_inner_outline_count(bars_a, bars_b) if include_inner_outline else 0
    return int(outer_n + inner_n)


# ---------------------------------------------------------------------------
# Orden de perímetro (horario, esquina superior-izquierda)
# ---------------------------------------------------------------------------

def _perimeter_ij_clockwise(nx, ny):
    if nx < 2 or ny < 2:
        return []
    out = []
    for ix in range(nx):
        out.append((ix, ny - 1))
    for iy in range(ny - 2, -1, -1):
        out.append((nx - 1, iy))
    for ix in range(nx - 2, -1, -1):
        out.append((ix, 0))
    for iy in range(1, ny - 1):
        out.append((0, iy))
    return out


def _outer_outline_ij_ordered(bars_a, bars_b):
    return _perimeter_ij_clockwise(bars_a, bars_b)


def _inner_outline_ij_ordered(bars_a, bars_b):
    nx = int(bars_a) - 2
    ny = int(bars_b) - 2
    if nx < 2 or ny < 2:
        return []
    return [
        (ix + 1, iy + 1)
        for ix, iy in _perimeter_ij_clockwise(nx, ny)
    ]


# ---------------------------------------------------------------------------
# Generación de puntos de rejilla
# ---------------------------------------------------------------------------

def generate_bar_points(
    center,
    side_short,
    side_long,
    short_on_x,
    bars_a,
    bars_b,
    cover,
    include_inner_outline,
    v_short=None,
    v_long=None,
):
    """
    Devuelve lista de ``{"pt": XYZ, "bar_enum": "A"|"B"|"IA"|"IB"}``.

    Si ``v_short`` y ``v_long`` son vectores unitarios, la rejilla se alinea
    a los ejes locales del símbolo; si son ``None``, se usa ±X/±Y proyecto.
    """
    offs_a = get_positions(side_short, bars_a, cover)
    offs_b = get_positions(side_long,  bars_b, cover)

    if len(offs_a) != bars_a or len(offs_b) != bars_b:
        raise Exception(u"Error interno en reparto de posiciones rejilla.")

    def pt_at(ix, iy):
        da = offs_a[ix]
        db = offs_b[iy]
        if v_short is not None and v_long is not None:
            try:
                dxy = v_short.Multiply(float(da)).Add(v_long.Multiply(float(db)))
                return center.Add(dxy)
            except Exception:
                pass
        if short_on_x:
            return XYZ(center.X + da, center.Y + db, center.Z)
        return XYZ(center.X + db, center.Y + da, center.Z)

    points = []
    for k, (ix, iy) in enumerate(_outer_outline_ij_ordered(bars_a, bars_b)):
        points.append(dict(pt=pt_at(ix, iy), bar_enum=("A" if (k % 2 == 0) else "B")))

    if not include_inner_outline or bars_a < 4 or bars_b < 4:
        return points

    for ik, (ix, iy) in enumerate(_inner_outline_ij_ordered(bars_a, bars_b)):
        points.append(dict(
            pt=pt_at(ix, iy),
            bar_enum=("IA" if ik % 2 == 0 else "IB"),
        ))
    return points


# ---------------------------------------------------------------------------
# Nombre de línea (alfabético) para Armadura_Ubicacion
# ---------------------------------------------------------------------------

def _linea_fierro_nombre_alfabetico(index):
    """0→A, 1→B, …, 25→Z, 26→AA, 27→AB, …"""
    index = max(0, int(index))
    out = []
    n = index
    while True:
        out.append(chr(ord("A") + (n % 26)))
        n = n // 26 - 1
        if n < 0:
            break
    return u"".join(reversed(out))


def _arma_len_mm_round_from_internal_ft(span_ft):
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertFromInternalUnits(float(span_ft), UnitTypeId.Millimeters)
    except Exception:
        return float(span_ft) * 304.8


# ---------------------------------------------------------------------------
# Fusión vertical de tramos
# ---------------------------------------------------------------------------

def fuse_vertical_world_intervals_from_jobs(
    jobs,
    short_curve_tolerance,
    xy_decimals=XY_KEY_DECIMALS_DEFAULT,
):
    """
    Fusiona tramos verticales que comparten la misma posición XY en el modelo.

    Devuelve lista de ``(base_xyz, span_z_ft, contrib_elem_ids, bar_enum_label)``.
    """
    tol = abs(float(short_curve_tolerance))
    if tol < 1e-12:
        tol = 1e-12
    decimals = int(xy_decimals)

    # Agrupar por (x_round, y_round, bar_enum)
    groups = defaultdict(list)
    for job in jobs or []:
        try:
            pts    = job["raw_pts"]
            height = float(job["height"])
            elem   = job.get("elem")
        except (KeyError, TypeError):
            continue
        for entry in pts or []:
            pt  = entry.get("pt")
            tag = entry.get("bar_enum", u"A")
            if pt is None:
                continue
            key = (round(float(pt.X), decimals), round(float(pt.Y), decimals), tag)
            groups[key].append((pt, height, elem))

    result = []
    for (rx, ry, tag), items in groups.items():
        if not items:
            continue
        # Fusión: barre los tramos de abajo hacia arriba y une los solapados
        items_sorted = sorted(items, key=lambda t: float(t[0].Z))
        fused_base = None
        fused_z_start = None
        fused_z_end   = None
        contrib_ids   = []
        base_z_base   = None

        def _flush():
            if fused_base is None:
                return None
            span = fused_z_end - fused_z_start
            if span < tol:
                return None
            return (fused_base, span, list(contrib_ids), tag)

        for pt, h, elem in items_sorted:
            z0 = float(pt.Z)
            z1 = z0 + h
            iv = _element_id_iv(elem) if elem is not None else -1
            if fused_base is None:
                fused_base   = pt
                fused_z_start = z0
                fused_z_end   = z1
                contrib_ids  = [iv] if iv >= 0 else []
            else:
                # Si este tramo empieza antes del final actual (o pegado), extender
                if z0 <= fused_z_end + tol:
                    if z1 > fused_z_end:
                        fused_z_end = z1
                    if iv >= 0 and iv not in contrib_ids:
                        contrib_ids.append(iv)
                else:
                    # Brecha → emitir y comenzar nuevo
                    seg = _flush()
                    if seg is not None:
                        result.append(seg)
                    fused_base    = pt
                    fused_z_start = z0
                    fused_z_end   = z1
                    contrib_ids   = [iv] if iv >= 0 else []

        seg = _flush()
        if seg is not None:
            result.append(seg)

    # Ordenar por Z de base luego por etiqueta para reproducibilidad
    result.sort(key=lambda r: (float(r[0].Z), r[3]))
    return result
