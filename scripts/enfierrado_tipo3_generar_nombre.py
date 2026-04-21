# -*- coding: utf-8 -*-
"""
Acción dedicada del botón "Generar Nombre" (Tipo 3).
"""

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog


def run(revit, inputs):
    nombre = (inputs.get("nombre") or "").strip()
    apellido = (inputs.get("apellido") or "").strip()

    full_name = u"{} {}".format(nombre, apellido).strip()
    if not full_name:
        TaskDialog.Show("Enfierrado - Tipo 3", u"Ingrese nombre y/o apellido.")
        return False

    TaskDialog.Show("Enfierrado - Tipo 3", u"Nombre generado:\n{}".format(full_name))
    return True

