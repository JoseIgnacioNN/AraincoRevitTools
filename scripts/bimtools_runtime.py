# -*- coding: utf-8 -*-
"""
Detección de runtime pyRevit / Revit (IronPython 2 vs CPython 3, 2024–2026).

Solo Revit 2024 + IronPython 2 soporta ProgressBar pyRevit, TransactionGroup
anidados e iteración ``GetOrderedParameters`` sin errores de expression trees.
"""

from __future__ import print_function

import sys

try:
    from bimtools_element_id import revit_version_year
except Exception:
    def revit_version_year(doc):
        return 0


def is_cpython3_runtime():
    """True si el intérprete es Python 3 (CPython o IronPython 3)."""
    try:
        return int(sys.version_info[0]) >= 3
    except Exception:
        return False


def is_ironpython2_runtime():
    """True solo en IronPython 2.x (pyRevit clásico en Revit 2024)."""
    try:
        if int(sys.version_info[0]) != 2:
            return False
        return u"IronPython" in str(sys.version)
    except Exception:
        return False


def revit_year_from_doc(doc):
    """Año Revit con respaldo ``VersionName`` / ``VersionBuild``."""
    year = 0
    try:
        year = int(revit_version_year(doc) or 0)
    except Exception:
        year = 0
    if year >= 2024:
        return year
    if doc is None:
        return year
    try:
        app = doc.Application
        for attr in (u"VersionName", u"VersionBuild", u"VersionNumber"):
            try:
                raw = getattr(app, attr, None)
            except Exception:
                raw = None
            if not raw:
                continue
            text = str(raw).replace(u",", u" ")
            for token in text.replace(u".", u" ").split():
                token = token.strip()
                if len(token) == 4 and token.isdigit():
                    y = int(token)
                    if y >= 2024:
                        return y
    except Exception:
        pass
    return year


def is_legacy_revit2024_armado(doc):
    """
    Entorno validado (Revit 2024, IronPython 2).

    Cualquier otro año o intérprete → comportamiento moderno (sin pbar/TG/scan).
    """
    if not is_ironpython2_runtime():
        return False
    try:
        year = int(revit_year_from_doc(doc))
    except Exception:
        year = 0
    return year == 2024


def use_rebar_setlayout_inclusion(doc=None):
    """``SetLayoutAs*`` en rebars: solo IronPython 2 + Revit 2024."""
    return is_legacy_revit2024_armado(doc)


def use_embed_solid_collision_probe(doc=None):
    """Prismas booleanos embed cabezal: solo legacy (2024); 2025+ omite sonda."""
    return is_legacy_revit2024_armado(doc)


def pyrevit_progress_bar_enabled(doc):
    return is_legacy_revit2024_armado(doc)


def use_transaction_group_armado_muros(doc, within_parent_transaction_group=False):
    if within_parent_transaction_group:
        return False
    return is_legacy_revit2024_armado(doc)


def skip_area_rein_ordered_parameters_scan(doc=None):
    if doc is not None:
        return not is_legacy_revit2024_armado(doc)
    return True
