"""
Script para Revit Python Shell (IronPython)
Obtiene todos los tipos de gancho de armadura (RebarHookType) del modelo
y los muestra en un TaskDialog.
"""

from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB.Structure import RebarHookType
from Autodesk.Revit.UI import TaskDialog

# doc es la variable predefinida del documento activo en Revit Python Shell
collector = FilteredElementCollector(doc)
hook_types = list(collector.OfClass(RebarHookType))
hook_names = [ht.Name for ht in hook_types]

# Preparar el mensaje para el TaskDialog
if hook_names:
    message = "\n".join(hook_names)
    title = "Tipos de gancho de armadura"
else:
    message = "No se encontraron tipos de gancho de armadura en el modelo."
    title = "Lista vacía"

TaskDialog.Show(title, message)
