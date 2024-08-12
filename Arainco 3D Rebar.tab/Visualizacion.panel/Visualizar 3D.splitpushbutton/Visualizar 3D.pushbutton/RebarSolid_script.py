# -*- coding: utf-8 -*-

# Importar los módulos necesarios
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, Transaction

# Obtener los servicios de Revit
from pyrevit import revit, DB
from pyrevit import script
__author__ = "[Tu Nombre]"

# Obtener el documento activo y la vista activa
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = uidoc.ActiveView

# Iniciar una transacción
with revit.Transaction("Set Rebar Solid and Unobscured In View"):
    # Recopilar todos los elementos de refuerzo
    rebars = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rebar)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    # Iterar sobre cada elemento de refuerzo
    for rebar in rebars:
        # Establecer el refuerzo como sólido y no oculto en la vista
        rebar.SetSolidInView(view, True)
        rebar.SetUnobscuredInView(view, True)


