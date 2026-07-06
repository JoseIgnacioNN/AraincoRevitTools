# -*- coding: utf-8 -*-
"""Zonas de estribos Ext/Cent por viga (regla 2·h)."""

from armado_vigas.domain.constants import ESTRIBO_INSET_MM, ESTRIBO_SPACING_DEFAULT_CENT, ESTRIBO_SPACING_DEFAULT_EXT, ESTRIBO_SPACING_MIN


def parse_beam_section(type_str):
    import re
    m = re.match(
        r"(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)",
        unicode(type_str or u"30×60"),
    )
    if m:
        return float(m.group(1)), float(m.group(2))
    return 30.0, 60.0


def section_height_mm(type_str):
    _, h = parse_beam_section(type_str)
    return int(round(h * 10.0))


def beam_array_length_mm(beam):
    return max(0, int(round(float(beam.get("len", 0.0)) * 1000.0) - 2 * ESTRIBO_INSET_MM))


def compute_stirrup_zones(beam):
    """
    Returns dict: mode ('triple'|'single'), zones list, singleKind optional.
    """
    h_mm = section_height_mm(beam.get("type"))
    l_arr = beam_array_length_mm(beam)
    sp_e = max(ESTRIBO_SPACING_MIN, int(beam.get("estExtSpacing") or ESTRIBO_SPACING_DEFAULT_EXT))
    sp_c = max(ESTRIBO_SPACING_MIN, int(beam.get("estCentSpacing") or ESTRIBO_SPACING_DEFAULT_CENT))
    min_len = max(sp_e, sp_c) * 0.2
    min_edge = 50.0
    lw = int(round(float(beam.get("len", 0.0)) * 1000.0))

    two_h = 2.0 * float(h_mm)
    if l_arr < two_h - 1e-6:
        return {
            "mode": "single",
            "singleKind": "cent",
            "L_ext_each": 0,
            "L_cent": l_arr,
            "zones": [{
                "role": "uni",
                "lenMm": l_arr,
                "fracStart": ESTRIBO_INSET_MM / float(lw) if lw else 0.0,
                "fracLen": l_arr / float(lw) if lw else 1.0,
            }],
        }

    l_ext_tgt = two_h
    l_half = 0.5 * l_arr
    l_ext_each = min(l_ext_tgt, l_half)
    l_cent = max(0.0, l_arr - 2.0 * l_ext_each)
    if l_cent < min_len + min_edge:
        return {
            "mode": "single",
            "singleKind": "merge",
            "L_ext_each": l_ext_each,
            "L_cent": l_cent,
            "zones": [{
                "role": "uni",
                "lenMm": l_arr,
                "fracStart": ESTRIBO_INSET_MM / float(lw) if lw else 0.0,
                "fracLen": l_arr / float(lw) if lw else 1.0,
            }],
        }

    s0 = ESTRIBO_INSET_MM
    zones = [
        {
            "role": "ext",
            "lenMm": int(round(l_ext_each)),
            "fracStart": s0 / float(lw) if lw else 0.0,
            "fracLen": l_ext_each / float(lw) if lw else 0.0,
        },
        {
            "role": "cent",
            "lenMm": int(round(l_cent)),
            "fracStart": (s0 + l_ext_each) / float(lw) if lw else 0.0,
            "fracLen": l_cent / float(lw) if lw else 0.0,
        },
        {
            "role": "ext",
            "lenMm": int(round(l_ext_each)),
            "fracStart": (s0 + l_ext_each + l_cent) / float(lw) if lw else 0.0,
            "fracLen": l_ext_each / float(lw) if lw else 0.0,
        },
    ]
    return {"mode": "triple", "L_ext_each": int(round(l_ext_each)), "L_cent": int(round(l_cent)), "zones": zones}
