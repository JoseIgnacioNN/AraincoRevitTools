# -*- coding: utf-8 -*-
"""
Módulo de prueba: Tipo 2 de formulario de enfierrado.

Este módulo se usa como "script dedicado por tipo" desde el formulario dinámico.
"""

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog


def run(revit, inputs):
    """Ejecuta la lógica de prueba del Tipo 2 (sin tocar la geometría)."""
    diametro = inputs.get("diametro", "")
    opcion = inputs.get("opcion", "")
    cantidad = inputs.get("cantidad", "")
    msg = u"Tipo 2 ejecutado.\nDiametro: {}\nOpcion: {}\nCantidad: {}".format(
        diametro, opcion, cantidad
    )
    TaskDialog.Show("Enfierrado - Tipo 2 (prueba)", msg)
    return True

