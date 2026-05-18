# -*- coding: utf-8 -*-
"""Operaciones puras sobre tramos verticales.

Estas funciones no importan Revit API; son el primer núcleo testeable del refactor.
"""

from column_reinforcement.models.segments import SegmentZ


def merge_sorted_cuts(cuts, merge_eps):
    """Une cortes cercanos manteniendo orden ascendente."""
    ordered = sorted(float(c) for c in cuts)
    merged = []
    eps = abs(float(merge_eps))
    for c in ordered:
        if not merged or c > merged[-1] + eps:
            merged.append(c)
    return merged


def split_z_span_by_cut_values(z_lo, z_hi, cut_values, tol):
    """Parte un tramo Z por cotas de corte ya calculadas."""
    z_lo = float(z_lo)
    z_hi = float(z_hi)
    tt = abs(float(tol))
    if tt < 1e-12:
        tt = 1e-12
    if z_hi <= z_lo + tt:
        return []

    valid = [z_lo, z_hi]
    for zi in cut_values or []:
        zc = float(zi)
        if zc > z_lo + tt and zc < z_hi - tt:
            valid.append(zc)

    merged = merge_sorted_cuts(valid, max(tt * 4.0, tt))
    if len(merged) < 2:
        return [SegmentZ(z_lo, z_hi - z_lo)]

    out = []
    for idx in range(len(merged) - 1):
        a = merged[idx]
        b = merged[idx + 1]
        dz = b - a
        if dz > tt * 4.0:
            out.append(SegmentZ(a, dz))
    if not out:
        return [SegmentZ(z_lo, z_hi - z_lo)]
    return out


def extend_intermediate_segments(segments, lap_length, eligible=True):
    """Alarga todos los segmentos salvo el último con la longitud de traslape."""
    out = []
    n = len(segments or [])
    lap = float(lap_length)
    for idx, seg in enumerate(segments or []):
        s = seg.copy() if hasattr(seg, "copy") else SegmentZ(seg[0], seg[1])
        if eligible and n > 1 and lap > 1e-12 and idx < n - 1:
            s.dz += lap
        out.append(s)
    return out


def apply_embedment_after_split(
    segments,
    top_lap_length,
    keep_top_embed,
    top_revoke_delta,
    keep_bottom_embed,
    bottom_revoke_delta,
    min_length,
):
    """Aplica empotramientos/revertidos al primer y último tramo después del troceo."""
    if not segments:
        return []

    out = [s.copy() if hasattr(s, "copy") else SegmentZ(s[0], s[1]) for s in segments]
    lap = float(top_lap_length)
    min_len = max(float(min_length), 0.0)

    if keep_top_embed and lap > 1e-12:
        out[-1].dz += lap
    elif float(top_revoke_delta) > 1e-12:
        out[-1].dz = max(out[-1].dz - float(top_revoke_delta), min_len)

    if keep_bottom_embed and lap > 1e-12:
        out[0].z_start -= lap
        out[0].dz += lap
    elif float(bottom_revoke_delta) > 1e-12:
        delta = float(bottom_revoke_delta)
        out[0].z_start += delta
        out[0].dz = max(out[0].dz - delta, min_len)

    return out
