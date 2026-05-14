# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from Autodesk.Revit import UI
from pyrevit import revit, forms
from Autodesk.Revit.DB import UnitUtils, UnitTypeId # Conversión de unidades
from Autodesk.Revit.Exceptions import OperationCanceledException
doc = revit.doc
uidoc = revit.uidoc

# ===================================================================================
# SELECCIONAR LA LOSA (HOST)
# ===================================================================================
try:
    ref = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element, "Selecciona una losa estructural")
    slab = doc.GetElement(ref)
except OperationCanceledException:
    forms.alert("No se seleccionó nada.", exitscript=True)

# ===================================================================================
# PROPIEDADES DE LA LOSA
# ===================================================================================
bbox = slab.get_BoundingBox(None) # Caja geométrica que encierra toda la losa
Z_inf = bbox.Min.Z # Cara inferior
Z_sup = bbox.Max.Z # Cara superior
e = slab.get_Parameter(DB.BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble() # Espesor
rec = UnitUtils.ConvertToInternalUnits(1.5, UnitTypeId.Centimeters) # Recubrimiento

# ===================================================================================
# INDICAR TIPO DE BARRA
# ===================================================================================
rebar_types = DB.FilteredElementCollector(doc).OfClass(DB.Structure.RebarBarType).WhereElementIsElementType().ToElements() # Lista de barras (Objeto)   
rebar_types_names = [rt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for rt in rebar_types] # Lista de barras (Nombre)
rebar_selected = forms.SelectFromList.show(rebar_types_names, title='Selecciona el diámetro de barra', button_name="Seleccionar") # Formulario con nombre de barras
rebar_selected = next((x for x, y in zip(rebar_types, rebar_types_names) if y == rebar_selected), None) # Encuentra el tipo de barra (objeto) asociado al nombre seleccionado
if not rebar_selected: # Control de error si cancela el formulario
    forms.alert("No se seleccionó ningún tipo de barra.", exitscript=True)
    
# ===================================================================================
# CONFIGURAR PLANO DE TRABAJO
# ===================================================================================
with revit.Transaction("Configurar Work Plane"):
    view = uidoc.ActiveView # Vista activa
    plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ.Zero) # Se configura un plano horizontal en el origen (Z=0)
    sketch_plane = DB.SketchPlane.Create(doc, plane) # Se crea el plano en el modelo
    view.SketchPlane = sketch_plane # Asigna el plano a la vista activa

# ===================================================================================
# ELEGIR TRAZADO DE BARRA
# ===================================================================================
snaps = UI.Selection.ObjectSnapTypes.Endpoints | UI.Selection.ObjectSnapTypes.Intersections | UI.Selection.ObjectSnapTypes.Nearest # Fuerza la configuración de snaps
try:
    picked_p1 = uidoc.Selection.PickPoint(snaps, "Paso 3: Haz clic en el PUNTO INICIAL") # Clic 1
    picked_p2 = uidoc.Selection.PickPoint(snaps, "Paso 4: Haz clic en el PUNTO FINAL") # Clic 2
except OperationCanceledException:
    forms.alert("Selección de puntos cancelada.", exitscript=True)

p1 = DB.XYZ(picked_p1.X, picked_p1.Y, Z_sup-rec)
p2 = DB.XYZ(picked_p2.X, picked_p2.Y, Z_sup-rec)
curves = [DB.Line.CreateBound(p1, p2)] # Las curvas deben estar en una lista

# ===================================================================================
# CREAR LA BARRA
# ===================================================================================
with revit.Transaction("Crear Barra"):
    try:
        rebar = DB.Structure.Rebar.CreateFromCurves(
            doc,
            DB.Structure.RebarStyle.Standard, # Estilo de barra (Standard o StirrupTie)
            rebar_selected, # Rebar type
            None, None, # Gancho inicial y final
            slab, # Anfitrión
            DB.XYZ(0, 0, 1), # Orientación de barra
            curves, # Curvas
            DB.Structure.RebarHookOrientation.Right, # Orientación de gancho inicial
            DB.Structure.RebarHookOrientation.Right, # Orientación de gancho final
            True, # Usar Rebar Shape existente
            True) # Si no existe Rebar Shape que coincida, Revit creará un Rebar Shape personalizado
        #print("Barra creada exitosamente ID: {}".format(rebar.Id))
        forms.alert("✅ Barra creada exitosamente. ID: " + rebar.Id.ToString(), title="Éxito")
    except Exception as e:
        print("Error al crear barra: {}".format(e))