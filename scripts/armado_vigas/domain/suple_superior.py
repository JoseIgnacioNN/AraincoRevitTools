# -*- coding: utf-8 -*-
"""Suple superior — por **apoyo** (columna/muro), L/3 en vigas adyacentes.

Fuente de verdad: ``session.suple_sup_apoyo_ids`` (ids de apoyo activos).

- Clic en apoyo del canvas / chips del rail → toggle.
- Por cada apoyo activo: vigas con ``colStart`` o ``colEnd`` = ese id reciben
  un tramo de longitud **L_viga / 3** desde ese extremo hacia el interior.
- Si dos vigas consecutivas comparten el apoyo (Fin+Ini), el tramo se fusiona
  en un solape continuo a través del nudo (L_a/3 + L_b/3).
- Losas (kind floor) **no** participan.

Los extremos ``start`` / ``end`` en specs siguen el orden del canvas
(izquierda→derecha). La colocación Revit traduce con ``axisReversed``.
"""

from __future__ import division

from armado_vigas.domain.constants import BAR_COUNT_MIN
from armado_vigas.domain.layers import beam_n_capas_sup, clamp_bar_count
from armado_vigas.domain.tramos import beams_share_bar_run_section

# Fracción del largo de viga que cubre el tramo de cada extremo (L/3).
SUPLE_END_PCT = 1.0 / 3.0
DEFAULT_DIAM_SUPLE_SUP_MM = 16
DEFAULT_N_SUPLE_SUP = 2

_FLOOR_KINDS = frozenset((u"floor", u"losa", u"slab"))


def ensure_beam_suple_superior(beam):
    """Inicializa campos de suple superior en el dict de viga (espejo derivado)."""
    if beam.get("supleSupEnabled") is None:
        beam["supleSupEnabled"] = False
    if beam.get("supleSupStartEnabled") is None:
        beam["supleSupStartEnabled"] = False
    if beam.get("supleSupEndEnabled") is None:
        beam["supleSupEndEnabled"] = False
    if beam.get("diamSupleSup") is None:
        beam["diamSupleSup"] = DEFAULT_DIAM_SUPLE_SUP_MM
    if beam.get("nSupleSup") is None:
        beam["nSupleSup"] = DEFAULT_N_SUPLE_SUP
    beam["nSupleSup"] = clamp_bar_count(beam["nSupleSup"])
    return beam


def ensure_session_suple_sup(session):
    """Campos de sesión para suple superior por apoyo."""
    if session is None:
        return None
    ids = getattr(session, u"suple_sup_apoyo_ids", None)
    if ids is None or not isinstance(ids, set):
        try:
            session.suple_sup_apoyo_ids = set(ids or [])
        except Exception:
            session.suple_sup_apoyo_ids = set()
    cfg = getattr(session, u"suple_sup_cfg_by_apoyo", None)
    if cfg is None or not isinstance(cfg, dict):
        try:
            session.suple_sup_cfg_by_apoyo = dict(cfg or {})
        except Exception:
            session.suple_sup_cfg_by_apoyo = {}
    if getattr(session, u"selected_suple_apoyo_id", None) is None:
        session.selected_suple_apoyo_id = None
    return session


def _normalize_arm_n(n, default=None):
    try:
        return clamp_bar_count(n)
    except Exception:
        return clamp_bar_count(default if default is not None else DEFAULT_N_SUPLE_SUP)


def _normalize_arm_diam(d, default=None):
    try:
        v = int(d)
        if v > 0:
            return v
    except Exception:
        pass
    return int(default if default is not None else DEFAULT_DIAM_SUPLE_SUP_MM)


def get_apoyo_suple_sup_arm(session, apoyo_id):
    """
    n · ø del suple SUP de un apoyo.

    Returns:
        ``(n_bars, diam_mm)``
    """
    ensure_session_suple_sup(session)
    aid = _as_unicode(apoyo_id or u"")
    if not aid or session is None:
        return DEFAULT_N_SUPLE_SUP, DEFAULT_DIAM_SUPLE_SUP_MM
    cfg = (session.suple_sup_cfg_by_apoyo or {}).get(aid) or {}
    n = _normalize_arm_n(cfg.get(u"n"), DEFAULT_N_SUPLE_SUP)
    diam = _normalize_arm_diam(cfg.get(u"diam"), DEFAULT_DIAM_SUPLE_SUP_MM)
    return n, diam


def set_apoyo_suple_sup_arm(session, apoyo_id, n=None, diam=None):
    """Actualiza n y/o ø del suple definido en ``apoyo_id``."""
    ensure_session_suple_sup(session)
    if session is None or not apoyo_id:
        return get_apoyo_suple_sup_arm(session, apoyo_id)
    aid = _as_unicode(apoyo_id)
    cfg = dict((session.suple_sup_cfg_by_apoyo or {}).get(aid) or {})
    if n is not None:
        cfg[u"n"] = _normalize_arm_n(n)
    else:
        cfg[u"n"] = _normalize_arm_n(cfg.get(u"n"), DEFAULT_N_SUPLE_SUP)
    if diam is not None:
        cfg[u"diam"] = _normalize_arm_diam(diam)
    else:
        cfg[u"diam"] = _normalize_arm_diam(cfg.get(u"diam"), DEFAULT_DIAM_SUPLE_SUP_MM)
    session.suple_sup_cfg_by_apoyo[aid] = cfg
    return cfg[u"n"], cfg[u"diam"]


def ensure_apoyo_suple_sup_arm_defaults(session, apoyo_id):
    """Crea entrada de armado si falta (al activar un apoyo)."""
    ensure_session_suple_sup(session)
    aid = _as_unicode(apoyo_id or u"")
    if not aid or session is None:
        return get_apoyo_suple_sup_arm(session, apoyo_id)
    existing = (session.suple_sup_cfg_by_apoyo or {}).get(aid)
    if existing is None or not isinstance(existing, dict):
        return set_apoyo_suple_sup_arm(
            session, aid, DEFAULT_N_SUPLE_SUP, DEFAULT_DIAM_SUPLE_SUP_MM
        )
    return get_apoyo_suple_sup_arm(session, aid)


def active_apoyos_suple_sup(session, apoyos=None):
    """Apoyos activos (ordenados) con su dict de catálogo si existe."""
    ensure_session_suple_sup(session)
    if session is None:
        return []
    ids = []
    try:
        ids = sorted(_as_unicode(x) for x in (session.suple_sup_apoyo_ids or set()) if x)
    except Exception:
        ids = []
    by_id = {}
    for ap in apoyos or getattr(session, u"apoyos", None) or []:
        if ap and ap.get(u"id"):
            by_id[_as_unicode(ap.get(u"id"))] = ap
    out = []
    for aid in ids:
        ap = by_id.get(aid) or {u"id": aid}
        n, diam = get_apoyo_suple_sup_arm(session, aid)
        out.append({u"apoyo": ap, u"id": aid, u"n": n, u"diam": diam})
    return out


def candidate_apoyos_suple_sup(session, apoyos=None):
    """Elegibles aún no activos (para añadir al set de suples)."""
    ensure_session_suple_sup(session)
    active = set()
    if session is not None:
        try:
            active = set(
                _as_unicode(x) for x in (session.suple_sup_apoyo_ids or set()) if x
            )
        except Exception:
            active = set()
    out = []
    for ap in eligible_apoyos_for_suple_sup(
        apoyos if apoyos is not None else getattr(session, u"apoyos", None)
    ):
        aid = _as_unicode(ap.get(u"id") or u"")
        if aid and aid not in active:
            out.append(ap)
    return out


def _as_unicode(v):
    try:
        return unicode(v)
    except NameError:
        return str(v)
    except Exception:
        try:
            return str(v)
        except Exception:
            return u""


def apoyo_kind(apoyo):
    if not apoyo:
        return u""
    try:
        return _as_unicode(apoyo.get("kind") or u"").lower()
    except Exception:
        return u""


def apoyo_is_floor(apoyo):
    """True si el apoyo es losa (excluida del suple SUP)."""
    if not apoyo:
        return False
    if apoyo_kind(apoyo) in _FLOOR_KINDS:
        return True
    try:
        aid = _as_unicode(apoyo.get("id") or u"")
        if aid.startswith(u"L") and not aid.startswith(u"C") and not aid.startswith(u"M"):
            # Prefijo L- de losa en adapters; no confundir con Marks raros.
            if apoyo_kind(apoyo) in _FLOOR_KINDS or not apoyo_kind(apoyo):
                return apoyo.get("element") is not None and apoyo_kind(apoyo) in _FLOOR_KINDS
    except Exception:
        pass
    return False


def apoyo_allows_suple_sup(apoyo):
    """Columnas y muros (no losas)."""
    if apoyo is None:
        return False
    if apoyo_is_floor(apoyo):
        return False
    k = apoyo_kind(apoyo)
    if k in (u"column", u"wall", u""):
        # kind vacío: si tiene id C-/M- se admite
        if k == u"":
            aid = _as_unicode(apoyo.get("id") or u"")
            return aid.startswith(u"C") or aid.startswith(u"M")
        return True
    return k not in _FLOOR_KINDS


def eligible_apoyos_for_suple_sup(apoyos):
    """Lista de apoyos (columnas/muros) sobre los que se puede definir suple."""
    out = []
    for ap in apoyos or []:
        if ap is None:
            continue
        if not apoyo_allows_suple_sup(ap):
            continue
        if not ap.get("id"):
            continue
        out.append(ap)
    return out


def is_apoyo_suple_sup_on(session, apoyo_id):
    ensure_session_suple_sup(session)
    if session is None or not apoyo_id:
        return False
    try:
        return _as_unicode(apoyo_id) in session.suple_sup_apoyo_ids
    except Exception:
        return False


def set_apoyo_suple_sup(session, apoyo_id, on, beams=None, apoyos=None):
    """Activa/desactiva suple en un apoyo y sincroniza vigas."""
    ensure_session_suple_sup(session)
    if session is None or not apoyo_id:
        return False
    aid = _as_unicode(apoyo_id)
    # Bloquear losas si vienen en el catálogo.
    if apoyos:
        for ap in apoyos:
            if ap and _as_unicode(ap.get("id")) == aid and not apoyo_allows_suple_sup(ap):
                return False
    if on:
        session.suple_sup_apoyo_ids.add(aid)
        ensure_apoyo_suple_sup_arm_defaults(session, aid)
        session.selected_suple_apoyo_id = aid
    else:
        session.suple_sup_apoyo_ids.discard(aid)
        if _as_unicode(getattr(session, u"selected_suple_apoyo_id", None) or u"") == aid:
            rest = sorted(session.suple_sup_apoyo_ids or set())
            session.selected_suple_apoyo_id = rest[0] if rest else None
    sync_beams_suple_from_apoyo_set(session, beams)
    sync_beams_suple_arm_from_apoyos(session, beams)
    return True


def toggle_apoyo_suple_sup(session, apoyo_id, beams=None, apoyos=None):
    """Toggle y retorna el nuevo estado (bool)."""
    on = not is_apoyo_suple_sup_on(session, apoyo_id)
    ok = set_apoyo_suple_sup(session, apoyo_id, on, beams=beams, apoyos=apoyos)
    if not ok:
        return False
    return is_apoyo_suple_sup_on(session, apoyo_id)


def select_apoyo_suple_sup(
    session,
    apoyo_id,
    beams=None,
    apoyos=None,
    activate_if_off=True,
):
    """
    Marca el apoyo como **activo de configuración** (n·ø del rail).

    Si ``activate_if_off`` y el apoyo aún no define suple, lo activa también.
    No desactiva un apoyo ya ON (eso es ``set_apoyo_suple_sup(..., False)``
    o toggle).

    Returns:
        ``True`` si queda seleccionado (y activo si se pidió activar).
    """
    ensure_session_suple_sup(session)
    if session is None or not apoyo_id:
        return False
    aid = _as_unicode(apoyo_id)
    if apoyos:
        for ap in apoyos:
            if ap and _as_unicode(ap.get("id")) == aid and not apoyo_allows_suple_sup(ap):
                return False
    if not is_apoyo_suple_sup_on(session, aid):
        if not activate_if_off:
            return False
        return bool(
            set_apoyo_suple_sup(
                session, aid, True, beams=beams, apoyos=apoyos
            )
        )
    ensure_apoyo_suple_sup_arm_defaults(session, aid)
    session.selected_suple_apoyo_id = aid
    return True


def clear_all_suple_sup_apoyos(session, beams=None):
    ensure_session_suple_sup(session)
    if session is None:
        return
    session.suple_sup_apoyo_ids = set()
    session.selected_suple_apoyo_id = None
    # Conservar cfg por si el usuario reactiva (no borrar armados).
    sync_beams_suple_from_apoyo_set(session, beams)


def resolve_suple_sup_arm_for_spec(session, spec, fallback_beam=None):
    """n · ø para un segmento de layout/guía a partir de su apoyo."""
    aid = None
    if isinstance(spec, dict):
        aid = spec.get(u"apoyo_id")
    n, diam = get_apoyo_suple_sup_arm(session, aid)
    if fallback_beam is not None and (not aid or n is None):
        try:
            ensure_beam_suple_superior(fallback_beam)
            if not aid:
                n = _normalize_arm_n(fallback_beam.get(u"nSupleSup"), n)
                diam = _normalize_arm_diam(fallback_beam.get(u"diamSupleSup"), diam)
        except Exception:
            pass
    return n, diam


def apply_apoyo_arm_to_adjacent_beams(session, apoyo_id, beams=None):
    """
    Espeja n/ø del apoyo a las vigas adyacentes (campos legados en dict de viga).

    Si una viga tiene dos extremos con apoyos distintos, prevalece el del
    extremo en escritura (última aplicación gana por lado vía campo único).
    """
    if beams is None and session is not None:
        beams = getattr(session, u"domain_beams", None)
    n, diam = get_apoyo_suple_sup_arm(session, apoyo_id)
    for rec in adjacent_beams_for_apoyo(beams, apoyo_id):
        beam = rec.get(u"beam")
        if beam is None:
            continue
        ensure_beam_suple_superior(beam)
        beam[u"nSupleSup"] = n
        beam[u"diamSupleSup"] = diam
    return n, diam


def sync_beams_suple_arm_from_apoyos(session, beams=None):
    """Propaga n/ø de cada apoyo activo a vigas colindantes (espejo legado)."""
    ensure_session_suple_sup(session)
    if beams is None and session is not None:
        beams = getattr(session, u"domain_beams", None)
    for aid in list(session.suple_sup_apoyo_ids or set()) if session else []:
        apply_apoyo_arm_to_adjacent_beams(session, aid, beams=beams)


def adjacent_beams_for_apoyo(beams, apoyo_id):
    """
    Vigas que tocan el apoyo en un extremo del canvas.

    Returns:
        Lista de ``{"beam", "index", "view_side", "len_m", "span_m"}`` con
        ``view_side`` en ``"start"`` | ``"end"`` y ``span_m = L/3``.
    """
    aid = _as_unicode(apoyo_id or u"")
    if not aid:
        return []
    out = []
    for i, beam in enumerate(beams or []):
        if beam is None:
            continue
        sides = []
        try:
            if _as_unicode(beam.get("colStart") or u"") == aid:
                sides.append(u"start")
            if _as_unicode(beam.get("colEnd") or u"") == aid:
                sides.append(u"end")
        except Exception:
            continue
        try:
            len_m = max(0.0, float(beam.get("len") or 0.0))
        except Exception:
            len_m = 0.0
        span_m = len_m * float(SUPLE_END_PCT)
        for side in sides:
            out.append({
                u"beam": beam,
                u"index": i,
                u"view_side": side,
                u"len_m": len_m,
                u"span_m": span_m,
                u"span_mm": int(round(span_m * 1000.0)),
            })
    return out


def sync_beams_suple_from_apoyo_set(session, beams=None):
    """
    Espeja ``suple_sup_apoyo_ids`` → flags por viga
    (``supleSupStartEnabled`` / ``End`` / master).
    """
    ensure_session_suple_sup(session)
    if beams is None and session is not None:
        beams = getattr(session, u"domain_beams", None)
    ids = set()
    if session is not None:
        try:
            ids = set(_as_unicode(x) for x in (session.suple_sup_apoyo_ids or set()))
        except Exception:
            ids = set()
    for beam in beams or []:
        ensure_beam_suple_superior(beam)
        cs = _as_unicode(beam.get("colStart") or u"")
        ce = _as_unicode(beam.get("colEnd") or u"")
        start_on = bool(cs and cs in ids)
        end_on = bool(ce and ce in ids)
        beam[u"supleSupStartEnabled"] = start_on
        beam[u"supleSupEndEnabled"] = end_on
        beam[u"supleSupEnabled"] = start_on or end_on
    return beams


def beam_suple_sup_enabled(beam):
    ensure_beam_suple_superior(beam)
    return bool(beam.get("supleSupEnabled"))


def beam_suple_sup_side_enabled(beam, view_side):
    """True si el suple superior está activo en el extremo canvas ``start`` o ``end``."""
    ensure_beam_suple_superior(beam)
    if not beam_suple_sup_enabled(beam):
        return False
    if view_side == "start":
        return bool(beam.get("supleSupStartEnabled"))
    if view_side == "end":
        return bool(beam.get("supleSupEndEnabled"))
    return False


def beam_suple_sup_layer_index(beam):
    """Índice 1-based: inmediatamente debajo de la última capa superior modelada."""
    return beam_n_capas_sup(beam) + 1


def suple_sup_metrics_mm(length_mm):
    """Longitud del tramo suple (L/3) en mm."""
    L = max(0, int(round(float(length_mm or 0))))
    span = int(round(L * SUPLE_END_PCT))
    return {
        "Lmm": L,
        "spanMm": span,
        "endPct": int(round(SUPLE_END_PCT * 100.0)),
    }


def beams_consecutive_for_suple(prev, cur):
    """Mismo criterio que corrida Tn superior en ``tramos.py``."""
    return beams_share_bar_run_section(prev, cur, es_cara_inferior=False)


def consecutive_pair_merges_suple(prev, cur):
    """
    Fusión: vigas consecutivas con suple activo en el extremo que las une
    (Fin de prev + Ini de cur), tipicamente el **mismo** apoyo central.
    """
    if not beams_consecutive_for_suple(prev, cur):
        return False
    if not (
        beam_suple_sup_enabled(prev)
        and beam_suple_sup_enabled(cur)
        and beam_suple_sup_side_enabled(prev, "end")
        and beam_suple_sup_side_enabled(cur, "start")
    ):
        return False
    # Mismo id de apoyo en el nudo (cuando está asignado).
    try:
        ae = _as_unicode(prev.get("colEnd") or u"")
        bs = _as_unicode(cur.get("colStart") or u"")
        if ae and bs and ae != bs:
            return False
    except Exception:
        pass
    return True


def beam_axis_reversed(beam):
    """True si LocationCurve 0 queda a la derecha del 1 en la vista activa."""
    return bool(beam and beam.get("axisReversed"))


def suple_sup_trim_from_curve_start(beam, view_side):
    """
    Mapea extremo en canvas (``start``=izquierda, ``end``=derecha) al extremo
    0/1 de LocationCurve para :func:`trim_line_end_portion`.
    """
    at_curve_start = view_side == "start"
    if beam_axis_reversed(beam):
        at_curve_start = not at_curve_start
    return at_curve_start


def suple_sup_resolver_at_view_side(beam, view_side):
    """
    Emp/gancho en el extremo libre según lado de canvas.

    Returns:
        ``(resolver_inicio, resolver_fin)`` — solo uno True.
    """
    rev = beam_axis_reversed(beam)
    if view_side == "start":
        return (not rev, rev)
    if view_side == "end":
        return (rev, not rev)
    return (False, False)


def merged_suple_sup_trim_sides(beam_a, beam_b):
    """
    Junta consecutiva (A izquierda, B derecha en canvas): extremos de curva
    que forman el tramo fusionado (cada uno L/3).
    """
    return (
        suple_sup_trim_from_curve_start(beam_a, "end"),
        suple_sup_trim_from_curve_start(beam_b, "start"),
    )


def trim_line_view_end_portion(line, beam, view_side, pct=None):
    """Recorta ``pct`` (default L/3) desde el extremo izquierdo o derecho del canvas."""
    if pct is None:
        pct = SUPLE_END_PCT
    return trim_line_end_portion(
        line,
        from_start=suple_sup_trim_from_curve_start(beam, view_side),
        pct=pct,
    )


def compute_suple_sup_segment_specs(sorted_beams, session=None):
    """
    Especificaciones de tramos suple superior en orden de vigas (``u``).

    Si hay ``session``, re-sincroniza flags desde apoyos antes de calcular.

    Returns:
        Lista de ``{"type": "start"|"end"|"merged", "indices": [...], "apoyo_id": ...}``.
    """
    if session is not None:
        sync_beams_suple_from_apoyo_set(session, sorted_beams)

    beams = list(sorted_beams or [])
    n = len(beams)
    specs = []

    for i in range(n):
        beam = beams[i]
        if not beam_suple_sup_enabled(beam):
            continue
        prev_merge = i > 0 and consecutive_pair_merges_suple(beams[i - 1], beam)
        next_merge = i < n - 1 and consecutive_pair_merges_suple(beam, beams[i + 1])
        if not prev_merge and beam_suple_sup_side_enabled(beam, "start"):
            specs.append({
                "type": "start",
                "indices": [i],
                "apoyo_id": beam.get("colStart"),
                "pct": SUPLE_END_PCT,
            })
        if not next_merge and beam_suple_sup_side_enabled(beam, "end"):
            specs.append({
                "type": "end",
                "indices": [i],
                "apoyo_id": beam.get("colEnd"),
                "pct": SUPLE_END_PCT,
            })

    for i in range(n - 1):
        if consecutive_pair_merges_suple(beams[i], beams[i + 1]):
            specs.append({
                "type": "merged",
                "indices": [i, i + 1],
                "apoyo_id": beams[i].get("colEnd") or beams[i + 1].get("colStart"),
                "pct": SUPLE_END_PCT,
            })

    return specs


def trim_line_fraction(line, frac_start, frac_end):
    """Recorta ``line`` entre fracciones normalizadas ``[0, 1]`` a lo largo del eje."""
    if line is None:
        return None
    try:
        from Autodesk.Revit.DB import Line

        fs = max(0.0, min(1.0, float(frac_start)))
        fe = max(0.0, min(1.0, float(frac_end)))
        if fe <= fs + 1e-9:
            return None
        L = float(line.Length)
        if L < 1e-9:
            return None
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        du = (p1 - p0).Normalize()
        pa = p0 + du.Multiply(L * fs)
        pb = p0 + du.Multiply(L * fe)
        if pa.DistanceTo(pb) < 1e-6:
            return None
        return Line.CreateBound(pa, pb)
    except Exception:
        return None


def trim_line_end_portion(line, from_start=False, pct=None):
    """Conserva ``pct`` de la longitud desde el extremo start o end (default L/3)."""
    if pct is None:
        pct = SUPLE_END_PCT
    if from_start:
        return trim_line_fraction(line, 0.0, pct)
    return trim_line_fraction(line, 1.0 - float(pct), 1.0)


def suple_sup_segments_layout_px(
    sorted_beams, layouts, content_w, pct_to_px_fn, session=None
):
    """
    Segmentos suple superior en px de alzado (L/3 por extremo o fusión L/3+L/3).
    """
    if session is not None:
        sync_beams_suple_from_apoyo_set(session, sorted_beams)
    if not beam_suple_sup_enabled_any(sorted_beams):
        return []

    specs = compute_suple_sup_segment_specs(sorted_beams, session=None)
    segs = []
    pct = float(SUPLE_END_PCT)
    for spec in specs:
        typ = spec.get("type")
        idxs = spec.get("indices") or []
        if typ == "merged" and len(idxs) >= 2:
            i, j = idxs[0], idxs[1]
            if i >= len(layouts) or j >= len(layouts):
                continue
            lay_a = layouts[i]
            lay_b = layouts[j]
            left_a = pct_to_px_fn(lay_a["leftPct"], content_w)
            width_a = pct_to_px_fn(lay_a["widthPct"], content_w)
            left_b = pct_to_px_fn(lay_b["leftPct"], content_w)
            width_b = pct_to_px_fn(lay_b["widthPct"], content_w)
            span_a = width_a * pct
            span_b = width_b * pct
            segs.append({
                "type": "merged",
                "indices": idxs,
                "x0": left_a + width_a - span_a,
                "x1": left_b + span_b,
                "junctionX": left_a + width_a,
                "merged": True,
                "apoyo_id": spec.get("apoyo_id"),
            })
        elif typ in ("start", "end") and idxs:
            i = idxs[0]
            if i >= len(layouts):
                continue
            lay_i = layouts[i]
            left = pct_to_px_fn(lay_i["leftPct"], content_w)
            width = pct_to_px_fn(lay_i["widthPct"], content_w)
            span_w = width * pct
            if typ == "start":
                x0, x1 = left, left + span_w
            else:
                x0, x1 = left + width - span_w, left + width
            segs.append({
                "type": typ,
                "indices": idxs,
                "x0": x0,
                "x1": x1,
                "merged": False,
                "apoyo_id": spec.get("apoyo_id"),
            })
    return sorted(segs, key=lambda s: s.get("x0", 0))


def beam_suple_sup_enabled_any(beams):
    for beam in beams or []:
        if (
            beam_suple_sup_side_enabled(beam, "start")
            or beam_suple_sup_side_enabled(beam, "end")
        ):
            return True
    return False
