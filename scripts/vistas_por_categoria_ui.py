# -*- coding: utf-8 -*-
"""
Arainco: Vistas por categoría — entrada UI.

Abre el formulario WPF y crea el conjunto 01_ENTREGABLE por categoría/zona.
Lógica en ``vistas_por_categoria/`` (service, constants, UI).

Revit 2024–2026 · IronPython (pyRevit).
"""

from __future__ import print_function

from vistas_por_categoria.ui.window import show_vistas_por_categoria_ui


def run(revit):
    """Punto de entrada desde el pushbutton."""
    show_vistas_por_categoria_ui(revit)


def main(revit):
    run(revit)
