# -*- coding: utf-8 -*-
"""Escenarios de confinamiento E (1ª capa) — port del mockup."""

from armado_vigas.domain.layers import first_layer_bar_count

# E(i–j) = estribo entre índices i y j; traba [k] = traba en índice k.
# Filtrado por first_layer_bar_count(beam) ∈ [2, 8].
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


def get_confin_scenarios(beam):
    n = first_layer_bar_count(beam)
    return CONFIN_BY_COUNT.get(n) or CONFIN_BY_COUNT[2]


def find_confin_def(beam):
    scenarios = get_confin_scenarios(beam)
    raw = beam.get("estConfin") or scenarios[0]["label"]
    for s in scenarios:
        if s["label"] == raw:
            return s
    return scenarios[0]


def ensure_beam_confinement(beam):
    opts = [s["label"] for s in get_confin_scenarios(beam)]
    cur = beam.get("estConfin")
    beam["estConfin"] = cur if cur in opts else opts[0]
    return beam["estConfin"]
