# -*- coding: utf-8 -*-
"""Entrada — Arainco: Mallas en muros.

Reutiliza el motor de Area Reinforcement de Armado Muros v3
(``armado_muros_lineales.run_mallas`` + ``armado_muros_preview_ui`` modo mallas
con shell elevación Machones / solo card Mallas).
"""

from __future__ import print_function


def run_pyrevit(uiapp):
    """Entrada canónica desde ``script.py`` tras el guardia de acceso."""
    from armado_muros_run import setup_armado_muros_paths, ensure_armado_muros_modules_fresh

    setup_armado_muros_paths()
    ensure_armado_muros_modules_fresh(force=True)
    from armado_muros_lineales import run_mallas

    return run_mallas(uiapp)


def run(uiapp):
    """Alias pyRevit / RPS."""
    return run_pyrevit(uiapp)
