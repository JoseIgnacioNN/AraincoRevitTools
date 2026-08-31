# -*- coding: utf-8 -*-
"""Confinamiento E (1ª capa): dibujo libre {perimetral, pairs, ties}.

Catálogo ``CONFIN_BY_COUNT`` queda solo como referencia/migración de labels
legacy (``estConfin``). La fuente de verdad operativa es ``estConfDraft``.
"""

from armado_vigas.domain.layers import first_layer_bar_count

# Labels de escenarios legacy (migración). E(i–j)=estribo; traba [k].
CONFIN_BY_COUNT = {
    2: [
        {"label": u"Perimetral", "perimetral": True, "pairs": [], "ties": []},
    ],
    3: [
        {"label": u"Perimetral", "perimetral": True, "pairs": [], "ties": []},
        {"label": u"Perimetral + traba [1]", "perimetral": True, "pairs": [], "ties": [1]},
    ],
    4: [
        {"label": u"Perimetral + traba [2]", "perimetral": True, "pairs": [], "ties": [2]},
        {"label": u"E(0–2) + E(1–3)", "perimetral": False, "pairs": [[0, 2], [1, 3]], "ties": []},
    ],
    5: [
        {"label": u"E(0–3) + E(1–4)", "perimetral": False, "pairs": [[0, 3], [1, 4]], "ties": []},
        {
            "label": u"E(0–3) + E(1–4) + traba [2]",
            "perimetral": False,
            "pairs": [[0, 3], [1, 4]],
            "ties": [2],
        },
    ],
    6: [
        {"label": u"E(0–3) + E(2–5)", "perimetral": False, "pairs": [[0, 3], [2, 5]], "ties": []},
        {
            "label": u"E(0–3) + E(2–5) + E(1–4)",
            "perimetral": False,
            "pairs": [[0, 3], [2, 5], [1, 4]],
            "ties": [],
        },
        {
            "label": u"E(0–3) + E(2–5) + traba [1] + traba [4]",
            "perimetral": False,
            "pairs": [[0, 3], [2, 5]],
            "ties": [1, 4],
        },
    ],
    7: [
        {"label": u"E(0–4) + E(2–6)", "perimetral": False, "pairs": [[0, 4], [2, 6]], "ties": []},
        {
            "label": u"E(0–4) + E(2–6) + traba [3]",
            "perimetral": False,
            "pairs": [[0, 4], [2, 6]],
            "ties": [3],
        },
        {
            "label": u"Perimetral + E(1–2) + E(4–5)",
            "perimetral": True,
            "pairs": [[1, 2], [4, 5]],
            "ties": [],
        },
        {
            "label": u"Perimetral + E(1–2) + E(4–5) + traba [3]",
            "perimetral": True,
            "pairs": [[1, 2], [4, 5]],
            "ties": [3],
        },
    ],
    8: [
        {
            "label": u"E(0–3) + E(2–5) + E(4–7)",
            "perimetral": False,
            "pairs": [[0, 3], [2, 5], [4, 7]],
            "ties": [],
        },
        {
            "label": u"E(0–3) + E(2–5) + E(4–7) + traba [1] + traba [6]",
            "perimetral": False,
            "pairs": [[0, 3], [2, 5], [4, 7]],
            "ties": [1, 6],
        },
    ],
}

FREEFORM_LABEL = u"Dibujo libre"


def empty_conf_draft():
    """Borrador por defecto: sin perimetral (usuario dibuja los confinamientos)."""
    return {u"perimetral": False, u"pairs": [], u"ties": []}


def _norm_pairs(pairs):
    out = []
    seen = set()
    for p in pairs or []:
        try:
            if not p or len(p) < 2:
                continue
            a, b = int(p[0]), int(p[1])
        except Exception:
            continue
        if a == b:
            continue
        key = (min(a, b), max(a, b))
        if key in seen:
            continue
        seen.add(key)
        out.append([key[0], key[1]])
    out.sort(key=lambda x: (x[0], x[1]))
    return out


def _norm_ties(ties):
    out = []
    seen = set()
    for t in ties or []:
        try:
            k = int(t)
        except Exception:
            continue
        if k in seen:
            continue
        seen.add(k)
        out.append(k)
    out.sort()
    return out


def normalize_conf_draft(draft, n_bars=None):
    if not isinstance(draft, dict):
        d = empty_conf_draft()
    else:
        d = {
            u"perimetral": bool(draft.get(u"perimetral")),
            u"pairs": _norm_pairs(draft.get(u"pairs")),
            u"ties": _norm_ties(draft.get(u"ties")),
        }
    if n_bars is not None:
        n = int(n_bars)
        d[u"pairs"] = [p for p in d[u"pairs"] if p[0] < n and p[1] < n]
        d[u"ties"] = [t for t in d[u"ties"] if t < n]
    return d


def conf_draft_signature(draft):
    d = normalize_conf_draft(draft)
    return (
        bool(d[u"perimetral"]),
        tuple(tuple(p) for p in d[u"pairs"]),
        tuple(d[u"ties"]),
    )


def get_confin_scenarios(beam):
    n = first_layer_bar_count(beam)
    return CONFIN_BY_COUNT.get(n) or CONFIN_BY_COUNT[2]


def conf_draft_label(draft, n_bars=None):
    """Etiqueta corta para UI; match catálogo si coincide exacto, si no dibujo libre."""
    d = normalize_conf_draft(draft, n_bars=n_bars)
    sig = conf_draft_signature(d)
    try:
        n = int(n_bars) if n_bars is not None else None
    except Exception:
        n = None
    scenarios = CONFIN_BY_COUNT.get(n) if n else None
    if scenarios:
        for sc in scenarios:
            if conf_draft_signature(sc) == sig:
                return sc.get(u"label") or FREEFORM_LABEL
    if not d[u"perimetral"] and not d[u"pairs"] and not d[u"ties"]:
        return FREEFORM_LABEL
    nE = len(d[u"pairs"])
    nT = len(d[u"ties"])
    parts = []
    if d[u"perimetral"]:
        parts.append(u"Peri")
    if nE:
        parts.append(u"{0}E".format(nE))
    if nT:
        parts.append(u"{0}T".format(nT))
    return u"+".join(parts) if parts else FREEFORM_LABEL


def is_conf_draft_defined(beam):
    """True si el usuario definió confinamiento (perimetral / estribos / trabas)."""
    d = get_conf_draft(beam)
    return bool(d.get(u"perimetral") or d.get(u"pairs") or d.get(u"ties"))


def get_conf_draft(beam):
    """Lee ``estConfDraft`` normalizado (sin mutar).

    Fuente de verdad: solo el borrador. No inventa estribos desde el catálogo
    legacy en lectura (evita pintar conf. en sección cuando aún no se definió).
    """
    if beam is None:
        return empty_conf_draft()
    n = first_layer_bar_count(beam)
    raw = beam.get(u"estConfDraft")
    if isinstance(raw, dict):
        return normalize_conf_draft(raw, n_bars=n)
    # Sin draft: vacío. Migración de label → ensure_beam_confinement (escritura).
    return empty_conf_draft()



def set_conf_draft(beam, draft):
    """Escribe borrador libre y sincroniza ``estConfin`` (etiqueta)."""
    if beam is None:
        return empty_conf_draft()
    n = first_layer_bar_count(beam)
    d = normalize_conf_draft(draft, n_bars=n)
    beam[u"estConfDraft"] = {
        u"perimetral": bool(d[u"perimetral"]),
        u"pairs": [list(p) for p in d[u"pairs"]],
        u"ties": list(d[u"ties"]),
    }
    beam[u"estConfin"] = conf_draft_label(d, n_bars=n)
    return d


def find_confin_def(beam):
    """Def. operativa para preview y Colocar: always from draft."""
    d = get_conf_draft(beam)
    n = first_layer_bar_count(beam) if beam is not None else 2
    return {
        u"label": conf_draft_label(d, n_bars=n),
        u"perimetral": bool(d[u"perimetral"]),
        u"pairs": [list(p) for p in d[u"pairs"]],
        u"ties": list(d[u"ties"]),
    }


def ensure_beam_confinement(beam):
    """Garantiza ``estConfDraft`` + ``estConfin`` coherentes.

    Idempotente: no reescribe el beam si ya está normalizado (evita set_conf_draft
    en cada repaint de hover).
    """
    if beam is None:
        return FREEFORM_LABEL
    n = first_layer_bar_count(beam)
    raw = beam.get(u"estConfDraft")
    if isinstance(raw, dict):
        d = normalize_conf_draft(raw, n_bars=n)
        label = conf_draft_label(d, n_bars=n)
        try:
            raw_pairs = raw.get(u"pairs") or []
            raw_ties = raw.get(u"ties") or []
            same = (
                bool(raw.get(u"perimetral")) == bool(d[u"perimetral"])
                and list(raw_pairs) == list(d[u"pairs"])
                and list(raw_ties) == list(d[u"ties"])
                and beam.get(u"estConfin") == label
            )
            if same:
                return label
        except Exception:
            pass
        set_conf_draft(beam, d)
        return beam[u"estConfin"]
    # Sin draft: vacío. No migrar catálogo legacy (eso pintaba E/T en sección
    # solo por label, sin que el usuario los hubiera dibujado).
    d = empty_conf_draft()
    set_conf_draft(beam, d)
    return beam[u"estConfin"]


def toggle_conf_estribo(beam, i0, i1):
    d = get_conf_draft(beam)
    pair = [min(int(i0), int(i1)), max(int(i0), int(i1))]
    if pair[0] == pair[1]:
        return d
    key = (pair[0], pair[1])
    pairs = [p for p in d[u"pairs"] if (p[0], p[1]) != key]
    if len(pairs) == len(d[u"pairs"]):
        pairs.append(pair)
    d[u"pairs"] = pairs
    return set_conf_draft(beam, d)


def toggle_conf_traba(beam, index):
    d = get_conf_draft(beam)
    k = int(index)
    ties = set(d[u"ties"])
    if k in ties:
        ties.discard(k)
    else:
        ties.add(k)
    d[u"ties"] = sorted(ties)
    return set_conf_draft(beam, d)


def toggle_conf_perimetral(beam):
    d = get_conf_draft(beam)
    d[u"perimetral"] = not bool(d[u"perimetral"])
    return set_conf_draft(beam, d)


def clear_conf_draft(beam):
    return set_conf_draft(beam, empty_conf_draft())
