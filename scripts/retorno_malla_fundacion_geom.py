# -*- coding: utf-8 -*-
"""Modos y constantes — Remate Mallas (pie fundación / superior coronamiento)."""

from __future__ import print_function

MODE_INFERIOR_FUND = u"inferior_fund"
MODE_SUPERIOR = u"superior"

MODE_OPTS = (
    (MODE_INFERIOR_FUND, u"Pie de muro · fundación corrida"),
    (MODE_SUPERIOR, u"Superior · coronamiento"),
)

COVER_SUPERIOR_MM = 25.0


def normalize_mode(mode):
    try:
        key = u"{0}".format(mode or u"").strip().lower()
    except Exception:
        key = u""
    aliases = {
        u"inferior": MODE_INFERIOR_FUND,
        u"pie": MODE_INFERIOR_FUND,
        u"fund": MODE_INFERIOR_FUND,
        u"fundacion": MODE_INFERIOR_FUND,
        u"inferior_fund": MODE_INFERIOR_FUND,
        u"superior": MODE_SUPERIOR,
        u"coronamiento": MODE_SUPERIOR,
        u"tope": MODE_SUPERIOR,
        u"cabezal": MODE_SUPERIOR,
    }
    return aliases.get(key, MODE_INFERIOR_FUND)


def mode_label(mode):
    m = normalize_mode(mode)
    for key, lab in MODE_OPTS:
        if key == m:
            return lab
    return MODE_OPTS[0][1]
