# -*- coding: utf-8 -*-
"""
Acción dedicada del botón "Hola Mario" (solo para Tipo 1).
"""

import clr

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog


def run(revit):
    TaskDialog.Show("Enfierrado - Hola Mario", u"Hola Mario")
    return True

