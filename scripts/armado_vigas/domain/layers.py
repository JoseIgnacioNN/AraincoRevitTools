# -*- coding: utf-8 -*-
"""Capas longitudinales sup/inf por viga."""

from armado_vigas.domain.constants import (
    BAR_COUNT_MIN,
    BAR_COUNT_MAX,
    CAPAS_DEFAULT,
    CAPAS_MAX,
    CAPAS_MIN,
)

# Cantidades y nº capas: compartidas entre todos los tramos Tn del lote.
LAYER_QTY_FIELDS = ("nSup", "nInf", "nSup2", "nInf2", "nSup3", "nInf3")
LAYER_CAPAS_FIELDS = ("nCapasSup", "nCapasInf")
GLOBAL_LAYER_SYNC_FIELDS = LAYER_QTY_FIELDS + LAYER_CAPAS_FIELDS


def layer_keys(layer_num):
    if layer_num == 1:
        return {
            "nSup": "nSup",
            "nInf": "nInf",
            "diamSup": "diamSup",
            "diamInf": "diamInf",
            "label": u"1ª",
        }
    return {
        "nSup": "nSup{0}".format(layer_num),
        "nInf": "nInf{0}".format(layer_num),
        "diamSup": "diamSup{0}".format(layer_num),
        "diamInf": "diamInf{0}".format(layer_num),
        "label": u"2ª" if layer_num == 2 else u"3ª",
    }


def clamp_bar_count(n):
    return max(BAR_COUNT_MIN, min(BAR_COUNT_MAX, int(round(n))))


def _clamp_capas(n):
    return max(CAPAS_MIN, min(CAPAS_MAX, int(round(n))))


def _sync_legacy_n_capas(beam):
    """Mantiene nCapas = max(sup, inf) para lectores heredados."""
    beam["nCapas"] = max(beam_n_capas_sup(beam), beam_n_capas_inf(beam))


def beam_n_capas_sup(beam):
    if beam.get("nCapasSup") is None:
        legacy = beam.get("nCapas")
        beam["nCapasSup"] = legacy if legacy is not None else CAPAS_DEFAULT
    return _clamp_capas(beam["nCapasSup"])


def beam_n_capas_inf(beam):
    if beam.get("nCapasInf") is None:
        legacy = beam.get("nCapas")
        beam["nCapasInf"] = legacy if legacy is not None else CAPAS_DEFAULT
    return _clamp_capas(beam["nCapasInf"])


def ensure_beam_layers(beam):
    n_sup = beam_n_capas_sup(beam)
    n_inf = beam_n_capas_inf(beam)
    beam["nCapasSup"] = n_sup
    beam["nCapasInf"] = n_inf
    _sync_legacy_n_capas(beam)
    for layer_num in range(2, CAPAS_MAX + 1):
        k = layer_keys(layer_num)
        if n_sup >= layer_num:
            if beam.get(k["nSup"]) is None:
                beam[k["nSup"]] = BAR_COUNT_MIN
            if beam.get(k["diamSup"]) is None:
                beam[k["diamSup"]] = beam.get("diamSup") or 16
        if n_inf >= layer_num:
            if beam.get(k["nInf"]) is None:
                beam[k["nInf"]] = BAR_COUNT_MIN
            if beam.get(k["diamInf"]) is None:
                beam[k["diamInf"]] = beam.get("diamInf") or 16
    sync_first_layer_bar_counts(beam)
    return beam["nCapas"]


def set_first_layer_bar_count(beam, count):
    """Cantidad 1ª capa sup/inf (ligadas — escenarios de confinamiento E)."""
    n = clamp_bar_count(count)
    beam["nSup"] = n
    beam["nInf"] = n
    return n


def sync_first_layer_bar_counts(beam):
    """Iguala nSup y nInf si difieren (datos heredados o desincronizados)."""
    ns = clamp_bar_count(beam.get("nSup") or BAR_COUNT_MIN)
    ni = clamp_bar_count(beam.get("nInf") or BAR_COUNT_MIN)
    if ns != ni:
        n = max(ns, ni)
        beam["nSup"] = n
        beam["nInf"] = n
    return beam["nSup"]


def first_layer_bar_count(beam):
    n = clamp_bar_count(beam.get("nSup") or beam.get("nInf") or BAR_COUNT_MIN)
    if n < 2:
        return 2
    return min(n, BAR_COUNT_MAX)


def is_global_layer_sync_field(field):
    return field in GLOBAL_LAYER_SYNC_FIELDS


def sync_layer_field_all_beams(beams, field, value):
    """
    Propaga cantidad de barras o nº capas (sup/inf) a todas las vigas del lote.
    La 1ª capa mantiene nSup = nInf (confinamiento E).
    """
    if not beams:
        return
    if field in ("nSup", "nInf"):
        for beam in beams:
            set_first_layer_bar_count(beam, value)
    elif field in GLOBAL_LAYER_SYNC_FIELDS:
        for beam in beams:
            beam[field] = value
    else:
        return
    from armado_vigas.domain.confinement import ensure_beam_confinement
    for beam in beams:
        ensure_beam_layers(beam)
        ensure_beam_confinement(beam)


def beam_layer_diam_sup(beam, layer_index_1based):
    k = layer_keys(layer_index_1based)
    return beam.get(k["diamSup"]) or beam.get("diamSup") or 16


def beam_layer_diam_inf(beam, layer_index_1based):
    k = layer_keys(layer_index_1based)
    return beam.get(k["diamInf"]) or beam.get("diamInf") or 16


def layer_bar_count(beam, layer_num, es_cara_inferior):
    k = layer_keys(layer_num)
    field = k["nInf"] if es_cara_inferior else k["nSup"]
    return clamp_bar_count(beam.get(field) or BAR_COUNT_MIN)
