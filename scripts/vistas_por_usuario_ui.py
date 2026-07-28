# -*- coding: utf-8 -*-
"""
Arainco: Vistas por usuario — entrada UI.

Abre el formulario WPF y crea el conjunto 02_TRABAJO por modelador.
Lógica en ``vistas_por_usuario/`` (service, people, UI).

Revit 2024–2026 · IronPython (pyRevit).
"""

from __future__ import print_function

from vistas_por_usuario.ui.window import show_vistas_por_usuario_ui


def run(revit):
    """Punto de entrada desde el pushbutton."""
    show_vistas_por_usuario_ui(revit)


def main(revit):
    run(revit)
