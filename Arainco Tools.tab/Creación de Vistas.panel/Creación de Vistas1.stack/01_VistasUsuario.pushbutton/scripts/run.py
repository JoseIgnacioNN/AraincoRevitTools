# -*- coding: utf-8 -*-
"""Punto de entrada lógico — Vistas por Usuario."""

from __future__ import print_function

from vistas_por_usuario.ui.window import show_vistas_por_usuario_ui


def main(revit_app):
    show_vistas_por_usuario_ui(revit_app)
