# -*- coding: utf-8 -*-
"""
Módulo de prueba: Tipo 1 de formulario de enfierrado.

Este módulo se usa como "script dedicado por tipo" desde el formulario dinámico.
"""

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog


def run(revit, inputs):
    """Ejecuta la lógica de prueba del Tipo 1 (sin tocar la geometría)."""
    campo_a = inputs.get("campo_a", "")
    campo_b = inputs.get("campo_b", "")
    msg = u"Tipo 1 ejecutado.\nCampo A: {}\nCampo B: {}".format(campo_a, campo_b)
    TaskDialog.Show("Enfierrado - Tipo 1 (prueba)", msg)
    return True

