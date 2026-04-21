# -*- coding: utf-8 -*-
"""
Acción dedicada del botón "Hola Jose" (solo para Tipo 2).
"""

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog


def run(revit):
    TaskDialog.Show("Enfierrado - Hola Jose", u"Hola Jose")
    return True

