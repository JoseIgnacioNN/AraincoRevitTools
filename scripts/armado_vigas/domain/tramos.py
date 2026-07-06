# -*- coding: utf-8 -*-
"""Tramos de barra Tn (corrida colineal + troceo por empalme)."""

from armado_vigas.domain.constants import MAX_BAR_MM
from armado_vigas.domain.stirrups import parse_beam_section

_SECTION_WIDTH_TOL_CM = 0.05

TRAMO_ACCENTS_SUP = (
    u"#22d3ee",
    u"#a78bfa",
    u"#34d399",
    u"#fbbf24",
    u"#818cf8",
    u"#2dd4bf",
)

TRAMO_ACCENTS_INF = (
    u"#fb7185",
    u"#e879f9",
    u"#f472b6",
    u"#fda4af",
    u"#c084fc",
    u"#f9a8d4",
)

# Alias heredado (p. ej. acentos sup en lectores antiguos).
TRAMO_ACCENTS = TRAMO_ACCENTS_SUP


def sort_beams(beams):
    return sorted(beams, key=lambda b: b.get("u", 0))


def beam_section_width_cm(beam):
    w_cm, _h_cm = parse_beam_section(beam.get("type"))
    return float(w_cm)


def beams_share_bar_run_section(prev, cur, es_cara_inferior=False):
    """
    Compatibilidad de sección para fusionar una corrida de barra.

    - **Superior** (``es_cara_inferior=False``): mismo ancho; la altura solo
      desplaza la fibra inferior (fondo de viga).
    - **Inferior** (``es_cara_inferior=True``): sección completa (ancho×alto).
    """
    if prev.get("colEnd") != cur.get("colStart"):
        return False
    if es_cara_inferior:
        return prev.get("type") == cur.get("type")
    w_prev = beam_section_width_cm(prev)
    w_cur = beam_section_width_cm(cur)
    return abs(w_prev - w_cur) <= _SECTION_WIDTH_TOL_CM


def continues_bar_run(sorted_beams, index, es_cara_inferior=False):
    if index <= 0:
        return False
    prev = sorted_beams[index - 1]
    cur = sorted_beams[index]
    return beams_share_bar_run_section(prev, cur, es_cara_inferior)


def build_bar_runs(sorted_beams, es_cara_inferior=False):
    runs = []
    current = None
    for i, beam in enumerate(sorted_beams):
        if current is None or not continues_bar_run(sorted_beams, i, es_cara_inferior):
            if current is not None:
                runs.append(current)
            current = {"section": beam.get("type"), "indices": [i]}
        else:
            current["indices"].append(i)
    if current is not None:
        runs.append(current)
    return runs


def compute_tramo_len_mm(sorted_beams, piece):
    idxs = piece.get("beamIndices") or []
    if not idxs:
        return 0
    mm = 0.0
    for k, i in enumerate(idxs):
        full = float(sorted_beams[i].get("len", 0.0)) * 1000.0
        if len(idxs) == 1:
            if piece.get("edgeStart") == "half" and piece.get("edgeEnd") == "half":
                return 0
            if piece.get("edgeStart") == "half" or piece.get("edgeEnd") == "half":
                mm += full * 0.5
            else:
                mm += full
            continue
        if k == 0 and piece.get("edgeStart") == "half":
            mm += full * 0.5
        elif k == len(idxs) - 1 and piece.get("edgeEnd") == "half":
            mm += full * 0.5
        else:
            mm += full
    return int(round(mm))


def subdivide_bar_run(sorted_beams, run, empalme_beam_ids, split_empalme):
    idxs = list(run.get("indices") or [])
    split_idxs = [
        i for i in idxs
        if split_empalme and sorted_beams[i].get("id") in empalme_beam_ids
    ]
    if not split_idxs:
        return [{
            "beamIndices": idxs,
            "section": run.get("section"),
            "edgeStart": None,
            "edgeEnd": None,
            "fromEmpalme": False,
        }]

    pieces = []
    run_pos = 0
    start_half = False
    for split_idx in split_idxs:
        pos = idxs.index(split_idx)
        seg = idxs[run_pos:pos + 1]
        if seg:
            pieces.append({
                "beamIndices": seg,
                "section": run.get("section"),
                "edgeStart": "half" if start_half else None,
                "edgeEnd": "half",
                "fromEmpalme": True,
            })
        run_pos = pos
        start_half = True
    tail = idxs[run_pos:]
    if tail:
        pieces.append({
            "beamIndices": tail,
            "section": run.get("section"),
            "edgeStart": "half" if start_half else None,
            "edgeEnd": None,
            "fromEmpalme": start_half,
        })
    return pieces


def build_tramos(
    sorted_beams,
    empalme_beam_ids=None,
    split_empalme=True,
    es_cara_inferior=False,
):
    empalme_beam_ids = empalme_beam_ids or set()
    out = []
    tid = 1
    for run in build_bar_runs(sorted_beams, es_cara_inferior=es_cara_inferior):
        for piece in subdivide_bar_run(
            sorted_beams, run, empalme_beam_ids, split_empalme
        ):
            ids = "+".join(sorted_beams[i].get("id", "?") for i in piece["beamIndices"])
            suffix = u" · traslapo" if piece.get("fromEmpalme") else u""
            out.append({
                "id": tid,
                "label": u"T{0} · {1}{2}".format(tid, ids, suffix),
                "section": piece.get("section"),
                "beamIndices": list(piece["beamIndices"]),
                "edgeStart": piece.get("edgeStart"),
                "edgeEnd": piece.get("edgeEnd"),
                "fromEmpalme": piece.get("fromEmpalme"),
                "inferredLenMm": compute_tramo_len_mm(sorted_beams, piece),
                "es_cara_inferior": bool(es_cara_inferior),
                "face": u"inf" if es_cara_inferior else u"sup",
            })
            tid += 1
    return out


def build_session_tramos(
    sorted_beams,
    empalme_beam_ids=None,
    empalme_beam_ids_sup=None,
    empalme_beam_ids_inf=None,
    split_empalme=True,
):
    """Devuelve ``(tramos_sup, tramos_inf)`` con acentos y metadatos de cara."""
    if empalme_beam_ids_sup is None and empalme_beam_ids_inf is None and empalme_beam_ids:
        empalme_beam_ids_sup = set(empalme_beam_ids)
        empalme_beam_ids_inf = set(empalme_beam_ids)
    empalme_sup = set(empalme_beam_ids_sup or [])
    empalme_inf = set(empalme_beam_ids_inf or [])
    tramos_sup = build_tramos(
        sorted_beams,
        empalme_sup,
        split_empalme,
        es_cara_inferior=False,
    )
    tramos_inf = build_tramos(
        sorted_beams,
        empalme_inf,
        split_empalme,
        es_cara_inferior=True,
    )
    add_tramo_accents(tramos_sup, es_cara_inferior=False)
    add_tramo_accents(tramos_inf, es_cara_inferior=True)
    return tramos_sup, tramos_inf


def tramo_exceeds_bar_limit(tramo):
    return int(tramo.get("inferredLenMm") or 0) > MAX_BAR_MM


def find_tramo_for_beam(tramos, beam_index):
    for t in tramos or []:
        if beam_index in (t.get("beamIndices") or []):
            return t
    return None


def find_tramo_half(tramos, beam_index, part):
    for t in tramos or []:
        idxs = t.get("beamIndices") or []
        if beam_index not in idxs:
            continue
        if part == 1:
            if idxs[-1] == beam_index and t.get("edgeEnd") == "half":
                return t
        elif idxs[0] == beam_index and t.get("edgeStart") == "half":
            return t
    return None


def add_tramo_accents(tramos, es_cara_inferior=False):
    accents = TRAMO_ACCENTS_INF if es_cara_inferior else TRAMO_ACCENTS_SUP
    for i, t in enumerate(tramos or []):
        t["accent"] = accents[i % len(accents)]
    return tramos


def format_dual_tramo_summary(sorted_beams, tramos_sup, tramos_inf):
    sup_txt = format_tramo_summary(sorted_beams, tramos_sup)
    inf_txt = format_tramo_summary(sorted_beams, tramos_inf)
    if sup_txt == u"—" and inf_txt == u"—":
        return u"—"
    return u"Sup: {0} · Inf: {1}".format(sup_txt, inf_txt)


def format_tramo_summary(sorted_beams, tramos):
    if not tramos:
        return u"—"
    parts = []
    for t in tramos:
        ids = u"+".join(sorted_beams[i].get("id", u"?") for i in (t.get("beamIndices") or []))
        parts.append(u"T{0}({1})".format(t.get("id"), ids))
    return u"{0} tramo(s) barra · {1}".format(len(tramos), u" · ".join(parts))
