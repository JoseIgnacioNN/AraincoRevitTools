# -*- coding: utf-8 -*-
__title__ = u"Show All"

import os
import sys

_pushbutton_dir = os.path.dirname(os.path.abspath(__file__))
if _pushbutton_dir not in sys.path:
    sys.path.insert(0, _pushbutton_dir)

import bimtools_access_bootstrap as _bimtools_access

if not _bimtools_access.require_tool_access(__file__, __revit__, __title__):
    raise SystemExit

from Autodesk.Revit import DB
from pyrevit import revit, forms
doc = revit.doc

view = doc.ActiveView
rebar_elements = DB.FilteredElementCollector(doc,view.Id).OfCategory(DB.BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements()
if rebar_elements:
    with revit.Transaction(u"Arainco: Show All"):
        for rebar in rebar_elements:
            try:
                rebar.SetPresentationMode(view, DB.Structure.RebarPresentationMode.All) # Cambiar el modo de presentación a 'Show All'
            except Exception as e:
                pass # Ignorar elementos que no soporten este cambio (ej. barras individuales)
else:
    forms.alert("No se encontraron barras de refuerzo visibles en la vista activa.", exitscript=True)