# -*- coding: utf-8 -*-
"""Punto de entrada lógico — Vistas por Categoría."""

from __future__ import print_function

from vistas_por_categoria.ui.window import show_vistas_por_categoria_ui


def main(revit_app):
    show_vistas_por_categoria_ui(revit_app)
