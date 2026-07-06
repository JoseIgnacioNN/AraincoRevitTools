# -*- coding: utf-8 -*-
"""Línea de fierro: huellas, etiquetas de despiece y parámetro Armadura_Ubicacion."""

from __future__ import print_function

import math

COLUMN_ARMA_UBICACION_PARAM = u"Armadura_Ubicacion"

DESPIECE_LIENZO_LEN_ROUND_MM = 10.0


def _arma_len_mm_round_from_internal_ft(length_ft):
    """Longitud en ft internas → mm redondeados (firma estable entre instancias)."""
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return round(
            float(
                UnitUtils.ConvertFromInternalUnits(
                    float(length_ft), UnitTypeId.Millimeters
                )
            ),
            3,
        )
    except Exception:
        try:
            return round(float(length_ft) * 304.8, 3)
        except Exception:
            return 0.0


def _linea_fierro_nombre_alfabetico(indice_cero):
    """0 → 'A', 25 → 'Z', 26 → 'AA'."""
    try:
        _chr = unichr
    except NameError:
        _chr = chr
    n = int(indice_cero) + 1
    out = u""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = _chr(65 + r) + out
    return out


def pata_flags_from_fingerprint(fp):
    try:
        k = int(fp[0])
    except Exception:
        return False, False
    if k == 1:
        return True, False
    if k == 2:
        return False, True
    if k == 3:
        return True, True
    return False, False


def fingerprint_seg_linea_fierro(span_seg_ft, ok_pat_bot, ok_pat_top, pata_len_ft):
    lz = _arma_len_mm_round_from_internal_ft(span_seg_ft)
    lp = _arma_len_mm_round_from_internal_ft(pata_len_ft)
    try:
        bot = bool(ok_pat_bot)
        top = bool(ok_pat_top)
    except Exception:
        bot = False
        top = False
    if bot and top:
        return (3, lz, lp, lp)
    if bot:
        return (1, lz, lp)
    if top:
        return (2, lz, lp)
    return (0, lz)


def linea_fierro_key_from_seg_jobs(seg_jobs_sorted, did_bot_list, did_top_list):
    fps = []
    for i, sj in enumerate(seg_jobs_sorted):
        db = did_bot_list[i] if i < len(did_bot_list) else bool(sj.get("want_bot_pata"))
        dt = did_top_list[i] if i < len(did_top_list) else bool(sj.get("want_top_pata"))
        fps.append(
            fingerprint_seg_linea_fierro(
                sj["span_seg"],
                db,
                dt,
                sj.get("pata_hook_ft_seg", 0.0),
            )
        )
    return (len(fps), tuple(fps))


def linea_fierro_sort_index_from_letter(letter):
    """Índice de orden A=0, B=1, … (para disposición del despiece izquierda→derecha)."""
    lab = (letter or u"").strip()
    if not lab:
        return 10 ** 6
    for i in range(10000):
        if _linea_fierro_nombre_alfabetico(i) == lab:
            return i
    return 10 ** 6


def linea_fierro_label_map_from_keys(keys):
    if not keys:
        return {}
    try:
        uniq = sorted(set(keys), key=lambda kk: (-kk[0], kk[1]))
    except Exception:
        return {}
    return {k: _linea_fierro_nombre_alfabetico(i) for i, k in enumerate(uniq)}


def align_fingerprints_to_seg_jobs(items, seg_jobs_sorted):
    """
    Huellas y patas del modelo indexadas por ``seg_i`` ascendente (orden del croquis).
    ``items``: entradas ``(troceo_ui_i, seg_i, rebar, fingerprint)``.
    """
    by_seg_i = {}
    for item in items or []:
        try:
            _troceo_ui, seg_i, _rb, fp = item
            by_seg_i[int(seg_i)] = fp
        except Exception:
            pass
    fps = []
    flags = []
    for sj in seg_jobs_sorted or []:
        si = int(sj["seg_i"])
        fp = by_seg_i.get(si)
        if fp is None:
            return None, None
        fps.append(fp)
        flags.append(pata_flags_from_fingerprint(fp))
    return tuple(fps), flags


def collect_linea_fierro_model_groups(line_plans, line_rb_accum):
    """
    Agrupa solo hilos con todos los tramos modelados (misma regla que
    ``Armadura_Ubicacion``). Devuelve asignaciones, mapa letra→clave y grupos
    ordenados para despiece.
    """
    assignments = []
    groups_by_key = {}
    for lp in line_plans or []:
        line_idx = lp.get("line_idx")
        n_seg_total = lp.get("n_seg_total")
        items = (line_rb_accum or {}).get(line_idx)
        if not items:
            continue
        items_sorted = sorted(items, key=lambda x: (x[0], x[1]))
        fps = [x[3] for x in items_sorted]
        rebars = [x[2] for x in items_sorted]
        if not fps or len(fps) != len(rebars) or len(fps) != n_seg_total:
            continue
        k = (len(fps), tuple(fps))
        for rb in rebars:
            assignments.append((rb, k))
        segs = sorted(lp.get("seg_jobs") or [], key=lambda s: int(s["seg_i"]))
        if not segs:
            continue
        fps_draw, pata_flags = align_fingerprints_to_seg_jobs(items, segs)
        if fps_draw is None:
            continue
        if k not in groups_by_key:
            groups_by_key[k] = {
                "key": k,
                "line_plan": lp,
                "seg_jobs": segs,
                "fingerprints": fps_draw,
                "pata_flags": pata_flags,
                "z_line_start_ft": float(lp.get("z_line_start", 0.0)),
            }
        else:
            existing_idx = groups_by_key[k]["line_plan"].get("line_idx", 10 ** 9)
            if lp.get("line_idx", 10 ** 9) < existing_idx:
                groups_by_key[k]["line_plan"] = lp
                groups_by_key[k]["seg_jobs"] = segs
                groups_by_key[k]["fingerprints"] = fps_draw
                groups_by_key[k]["pata_flags"] = pata_flags
                groups_by_key[k]["z_line_start_ft"] = float(
                    lp.get("z_line_start", 0.0)
                )
    label_map = linea_fierro_label_map_from_keys(list(groups_by_key.keys()))
    groups_ordered = sorted(
        groups_by_key.values(),
        key=lambda g: linea_fierro_sort_index_from_letter(label_map.get(g["key"])),
    )
    return assignments, label_map, groups_ordered


def apply_linea_fierro_armadura_ubicacion(assignments_list):
    if not assignments_list:
        return
    nombre_por_key = linea_fierro_label_map_from_keys([k for _rb, k in assignments_list])
    if not nombre_por_key:
        return
    for rb, k in assignments_list:
        try:
            nm = nombre_por_key.get(k)
            if nm is None or rb is None:
                continue
            p = rb.LookupParameter(COLUMN_ARMA_UBICACION_PARAM)
            if p is None or p.IsReadOnly:
                continue
            p.Set(u"{}".format(nm))
        except Exception:
            continue


DIAMETER_SYMBOL = u"\u00f8"


def _despiece_diameter_text(diameter_mm):
    try:
        d = int(round(float(diameter_mm)))
    except Exception:
        d = 0
    return u"{0}{1}".format(DIAMETER_SYMBOL, d)


def _despiece_parcial_mm_ceiling_10(mm):
    try:
        v = float(mm)
    except Exception:
        return 0
    if v <= 1e-9:
        return 0
    step = float(DESPIECE_LIENZO_LEN_ROUND_MM)
    return int(math.ceil(v / step) * step)


def etiqueta_despiece_mm(diameter_mm, fp):
    """
    Etiqueta de despiece en lienzo: parciales con ceiling a 10 mm; L total = suma de parciales.
    """
    dia = _despiece_diameter_text(diameter_mm)
    kind = fp[0]
    if kind == 0:
        lz = _despiece_parcial_mm_ceiling_10(fp[1])
        return u"{0} L={1}".format(dia, lz)
    if kind == 1:
        lz = _despiece_parcial_mm_ceiling_10(fp[1])
        lp = _despiece_parcial_mm_ceiling_10(fp[2])
        total = lz + lp
        return u"{0} L={1} ({2}+{3})".format(dia, total, lp, lz)
    if kind == 2:
        lz = _despiece_parcial_mm_ceiling_10(fp[1])
        lp = _despiece_parcial_mm_ceiling_10(fp[2])
        total = lz + lp
        return u"{0} L={1} ({2}+{3})".format(dia, total, lz, lp)
    if kind == 3:
        lz = _despiece_parcial_mm_ceiling_10(fp[1])
        lp1 = _despiece_parcial_mm_ceiling_10(fp[2])
        lp2 = _despiece_parcial_mm_ceiling_10(fp[3])
        total = lz + lp1 + lp2
        return u"{0} L={1} ({2}+{3}+{4})".format(dia, total, lp1, lz, lp2)
    return u"{0} L=0".format(dia)


def etiqueta_empalme_mm(lap_mm):
    """Distancia de empalme/traslape en la zona de solape del despiece, p. ej. ``(860)``."""
    try:
        v = int(round(float(lap_mm)))
    except Exception:
        v = 0
    return u"({0})".format(v)
