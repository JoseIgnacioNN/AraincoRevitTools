# -*- coding: utf-8 -*-
"""
Arainco: Láminas por categoría — entrada UI.

Abre el formulario WPF y crea ViewSheet con correlativo por Clasificacion.
Lógica en ``laminas_por_categoria/`` (service, constants, UI).

Revit 2024–2026 · IronPython (pyRevit).
"""

from __future__ import print_function

from laminas_por_categoria.ui.window import show_laminas_por_categoria_ui


def run(revit):
    """Punto de entrada desde el pushbutton."""
    show_laminas_por_categoria_ui(revit)


def main(revit):
    run(revit)
