# -*- coding: utf-8 -*-
"""Reglas puras para etiquetas de esquema de barras."""


def is_lap_extension_scheme(tag):
    """Esquemas que reciben traslape en tramos intermedios."""
    return str(tag or "").strip() in ("A", "IA", "B", "IB")


def is_a_split_scheme(tag):
    """Esquemas tipo A (corners / eje corto): mismo troceo que A en planos de referencia sin desplazamiento."""
    return str(tag or "").strip() in ("A", "IA")


def is_b_split_scheme(tag):
    """Esquemas que usan planos B desplazados."""
    return str(tag or "").strip() in ("B", "IB")
