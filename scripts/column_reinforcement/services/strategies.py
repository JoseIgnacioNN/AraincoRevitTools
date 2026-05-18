# -*- coding: utf-8 -*-
"""Strategies puras/semipuras para políticas de armado."""

from column_reinforcement.geometry.schemes import (
    is_a_split_scheme,
    is_b_split_scheme,
    is_lap_extension_scheme,
)
from column_reinforcement.geometry.segments import extend_intermediate_segments


class SplitPolicyA(object):
    """Política A / IA: planos base sin desplazamiento."""

    def applies_to(self, scheme_tag):
        return is_a_split_scheme(scheme_tag)


class SplitPolicyB(object):
    """Política B/IB: planos desplazados por L(Ø) según eje de columna."""

    def applies_to(self, scheme_tag):
        return is_b_split_scheme(scheme_tag)


class LapExtensionPolicy(object):
    """Traslape tabular para todos los tramos salvo el último."""

    def applies_to(self, scheme_tag):
        return is_lap_extension_scheme(scheme_tag)

    def apply(self, segments, lap_length, scheme_tag):
        return extend_intermediate_segments(
            segments,
            lap_length,
            eligible=self.applies_to(scheme_tag),
        )
