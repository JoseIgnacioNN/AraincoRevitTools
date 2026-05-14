# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from Autodesk.Revit import UI
from Autodesk.Revit.DB import UnitUtils, UnitTypeId
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, forms
from System.Collections.Generic import List
import math
doc = revit.doc
uidoc = revit.uidoc
grupo_transacciones = DB.TransactionGroup(doc, "ARAINCO - Losa Armadura Superior") # Fusiona todas las transacciones en una sola
grupo_transacciones.Start()

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
Z_inf = bbox.Min.Z # Elevación de cara inferior
Z_sup = bbox.Max.Z # Elevación de cara superior
e = slab.get_Parameter(DB.BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble() # Espesor

# Recubrimiento superior
rec_top_param = slab.get_Parameter(DB.BuiltInParameter.CLEAR_COVER_TOP) # Parámetro del recubrimiento superior
rec_top_id = rec_top_param.AsElementId() # Id del tipo de recubrimiento
rec_top = doc.GetElement(rec_top_id).CoverDistance # Recubrimiento

# ===================================================================================
# INDICAR TIPO DE BARRA
# ===================================================================================
rebar_types = DB.FilteredElementCollector(doc).OfClass(DB.Structure.RebarBarType).WhereElementIsElementType().ToElements() # Lista de barras (Objeto)   
rebar_types_names = [rt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for rt in rebar_types] # Lista de barras (Nombre)
rebar_selected = forms.SelectFromList.show(rebar_types_names, title='Selecciona el diámetro de barra', button_name="Seleccionar") # Formulario con nombre de barras
rebar_selected = next((x for x, y in zip(rebar_types, rebar_types_names) if y == rebar_selected), None) # Encuentra el tipo de barra (objeto) asociado al nombre seleccionado

if not rebar_selected: # Control de error si cancela el formulario
    forms.alert("No se seleccionó ningún tipo de barra.", exitscript=True)

rebar_fi = rebar_selected.get_Parameter(DB.BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble() # Diámetro seleccionado

# ===================================================================================
# INDICAR ESPACIAMIENTO
# ===================================================================================
esp = forms.ask_for_string(prompt="Ingresa el espaciamiento en centímetros (ej. 20):", title="Espaciamiento de la Armadura")
if esp>0:
    esp = UnitUtils.ConvertToInternalUnits(float(esp), UnitTypeId.Centimeters)
if not esp:
    forms.alert("Operación cancelada por el usuario.", exitscript=True)

# ===================================================================================
# CONFIGURAR PLANO DE TRABAJO
# ===================================================================================
with revit.Transaction("ARAINCO - Configurar Work Plane"):
    view = uidoc.ActiveView # Vista activa
    plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ.Zero) # Se configura un plano horizontal en el origen (Z=0)
    sketch_plane = DB.SketchPlane.Create(doc, plane) # Se crea el plano en el modelo
    view.SketchPlane = sketch_plane # Asigna el plano a la vista activa

# ===================================================================================
# BUSCAR FAMILIA DE LÍNEA
# ===================================================================================
nombre_familia = "EST_D_DEATIL ITEM_EMPALME"
tipos = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements() # Tipos de familia (Family Symbols) cargados en el proyecto
tipo = next((tip for tip in tipos if hasattr(tip, "Family") and tip.Family.Name == nombre_familia), None) # Se encuentra el tipo que coincide con el nombre indicado

if not tipo:
    forms.alert("No se encontró la familia '{}' en el proyecto.\nPor favor, cárgala antes de usar esta herramienta.".format(nombre_familia), exitscript=True)

# ===================================================================================
# FUNCIÓN PARA DIBUJAR CON RUBBER BAND
# ===================================================================================
def seleccionar_puntos():
    id_capturado = []

    # Se crea una función que captura los datos de la línea que se va a crear
    def capturador(sender, args):
        nuevos_ids = args.GetAddedElementIds()
        if nuevos_ids.Count > 0:
            id_capturado.append(nuevos_ids[0]) # Captura el ID de la línea
    doc.Application.DocumentChanged += capturador # Se activa la función en segundo plano

    # Crear línea
    try:
        uidoc.PromptForFamilyInstancePlacement(tipo) # Herramienta de dibujo nativa (con rubber band)
    except OperationCanceledException:
        pass # Avanzar con el código si el usuario presionó ESC para terminar
    finally:
        doc.Application.DocumentChanged -= capturador # Se apaga la función en segundo plano

    if not id_capturado:
        forms.alert("Operación cancelada. No dibujaste la línea.", exitscript=True)

    # Se extraen las coordenadas y se borra la línea temporal
    with revit.Transaction("ARAINCO - Selección de puntos"):
        linea_id = id_capturado[0] # Id de la primera línea dibujada
        p1 = doc.GetElement(id_capturado[0]).Location.Curve.GetEndPoint(0)
        p2 = doc.GetElement(id_capturado[0]).Location.Curve.GetEndPoint(1)
        doc.Delete(linea_id) # Se borra la línea temporal
    return p1, p2

# ===================================================================================
# TRAZADO DE BARRA
# ===================================================================================

# Seleccionar largo de vanos
p1, p2 = seleccionar_puntos()
L1 = ((p2.X-p1.X)**2+(p2.Y-p1.Y)**2)**0.5/3
p1, p2 = seleccionar_puntos()
L2 = ((p2.X-p1.X)**2+(p2.Y-p1.Y)**2)**0.5/3

# Seleccionar recorrido
p1, p2 = seleccionar_puntos()
L_recorrido = ((p2.X-p1.X)**2+(p2.Y-p1.Y)**2)**0.5
ang = math.atan2(p2.Y-p1.Y,p2.X-p1.X)
vdir = (p2 - p1).Normalize() # Vector director normalizado (formato requerido por la API)

# Geometría de la barra
Z = Z_sup - rec_top - rebar_fi/2
p_start = DB.XYZ(p1.X-L1*math.sin(ang), p1.Y+L1*math.cos(ang), Z) # Punto inicial de la barra
p_end = DB.XYZ(p1.X+L2*math.sin(ang), p1.Y-L2*math.cos(ang), Z) # Punto final de la barra
curves = [DB.Line.CreateBound(p_start, p_end)] # Las puntos deben estar en una lista

# ===================================================================================
# CREAR LA BARRA
# ===================================================================================

# Barra individual
with revit.Transaction("ARAINCO - Crear Rebar Set"):
    try:
        rebar = DB.Structure.Rebar.CreateFromCurves(
            doc,
            DB.Structure.RebarStyle.Standard, # Estilo de barra (Standard o StirrupTie)
            rebar_selected, # Rebar type
            None, None, # Gancho inicial y final
            slab, # Anfitrión
            vdir, # Vector director
            curves, # Curvas
            DB.Structure.RebarHookOrientation.Right, # Orientación de gancho inicial
            DB.Structure.RebarHookOrientation.Right, # Orientación de gancho final
            True,True) # Usar Rebar Shape existente. Si no existe Rebar Shape que coincida, Revit creará un Rebar Shape personalizado

# Transformar a Rebar Set
        N_barras = math.floor(L_recorrido/esp+1)
        rebar.GetShapeDrivenAccessor().SetLayoutAsNumberWithSpacing(
            N_barras,   # Número de barras
            esp,  # Espaciamiento
            True, True, True)   # True = Distribuye hacia la derecha de la línea. Incluir barra inicial. Incluir barra final.
        rebar.SetPresentationMode(uidoc.ActiveView, DB.Structure.RebarPresentationMode.Middle)

# Multi-Rebar Annotation
        nombre_tipo = "Recorrido Barras"
        tipos = DB.FilteredElementCollector(doc).OfClass(DB.MultiReferenceAnnotationType)
        tipo = next((t for t in tipos if DB.Element.Name.GetValue(t) == nombre_tipo), None) # Se encuentra la cota indicada
        
        if tipo:
            tipo_opts = DB.MultiReferenceAnnotationOptions(tipo) # Se abre la configuración de Multi-Rebar Annotation
            rebar_ids = List[DB.ElementId]()
            rebar_ids.Add(rebar.Id)
            tipo_opts.SetElementsToDimension(rebar_ids)
            tipo_opts.DimensionLineDirection = vdir
            tipo_opts.DimensionPlaneNormal = uidoc.ActiveView.ViewDirection
            offset = UnitUtils.ConvertToInternalUnits(80, UnitTypeId.Centimeters) # Desfase de la cota respecto al centro de la barra
            mid_point = DB.XYZ((p1.X + p2.X)/2-offset*math.sin(ang), (p1.Y + p2.Y)/2+offset*math.cos(ang), curves[0].GetEndPoint(0).Z) # Se centra la cota al punto medio del recorrido y luego se suma el offset
            tipo_opts.DimensionLineOrigin = mid_point
            tipo_opts.TagHeadPosition = mid_point
            DB.MultiReferenceAnnotation.Create(doc, uidoc.ActiveView.Id, tipo_opts)
            forms.alert("✅ Rebar Set y Anotación creados exitosamente. ID: " + rebar.Id.ToString(), title="Éxito")
        else:
            forms.alert("La armadura se creó, pero no se encontró el tipo de cota llamado '{}'.\n\nRevisa que esté cargado en el proyecto o que el nombre esté escrito exactamente igual.".format(nombre_tipo), title="Falta Tipo de Cota")
        
    except Exception as e:
        print("Error al crear barra o anotación: {}".format(e))

grupo_transacciones.Assimilate()
