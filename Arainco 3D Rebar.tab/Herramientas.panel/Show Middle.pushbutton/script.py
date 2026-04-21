# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from pyrevit import revit, forms
doc = revit.doc

view = doc.ActiveView
rebar_elements = DB.FilteredElementCollector(doc,view.Id).OfCategory(DB.BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
if rebar_elements:
    with revit.Transaction('ARAINCO - Show All'):
        for rebar in rebar_elements:
            try:
                rebar.SetPresentationMode(view, DB.Structure.RebarPresentationMode.Middle) # Cambiar el modo de presentación a 'Middle'
            except Exception as e:
                pass # Ignorar elementos que no soporten este cambio (ej. barras individuales)
else:
    forms.alert("No se encontraron barras de refuerzo visibles en la vista activa.", exitscript=True)