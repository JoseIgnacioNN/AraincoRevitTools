# -*- coding: utf-8 -*-
"""
Cuantificación de armadura en losas — tablas por nivel / malla / ubicación.

Abre UI WPF para seleccionar niveles (checkbox) y generar las tablas.

Revit 2024+ | pyRevit / IronPython.
"""

from __future__ import print_function

from cuantificacion_losa_ui import mostrar_aviso, show_cuantificacion_losa_window

__title__ = u"Arainco: Cuantificación losa por nivel"


def run(uiapp):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        mostrar_aviso(uiapp, u"No hay documento activo.")
        return

    doc = uidoc.Document
    if doc.IsFamilyDocument:
        mostrar_aviso(
            uiapp,
            u"Esta herramienta solo funciona en un proyecto (no en familias).",
        )
        return

    show_cuantificacion_losa_window(uiapp)
