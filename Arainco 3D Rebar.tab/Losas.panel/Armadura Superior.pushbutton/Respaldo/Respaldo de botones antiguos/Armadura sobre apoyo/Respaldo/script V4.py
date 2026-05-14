# -*- coding: utf-8 -*-
import math
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
rebar_fi = float(rebar_selected[1:]) # Diámetro seleccionado
rebar_selected = next((x for x, y in zip(rebar_types, rebar_types_names) if y == rebar_selected), None) # Encuentra el tipo de barra (objeto) asociado al nombre seleccionado
if not rebar_selected: # Control de error si cancela el formulario
    forms.alert("No se seleccionó ningún tipo de barra.", exitscript=True)

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
tipo = None
tipos = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements() # Tipos de familia (Family Symbols) cargados en el proyecto

for tip in tipos:
    if hasattr(tip, "Family") and tip.Family.Name == nombre_familia: # Se encuentra el tipo que coincide con el nombre indicado
        tipo = tip
        break # Apenas encuentra el primer tipo de esta familia, se detiene

if not tipo:
    forms.alert("No se encontró la familia '{}' en el proyecto.\nPor favor, cárgala antes de usar esta herramienta.".format(nombre_familia), exitscript=True)

# ===================================================================================
# FUNCIÓN PARA DIBUJAR CON RUBBER BAND
# ===================================================================================
def obtener_puntos_con_lapiz(mensaje_alerta):
    forms.alert("{}\n\nINSTRUCCIONES:\n1. Dibuja la línea (verás el efecto elástico).\n2. Presiona 'ESC' una vez para continuar al siguiente paso.".format(mensaje_alerta))
    
    # Capturamos el ID máximo antes de dibujar para saber qué elemento es el nuevo
    max_id_before = max([x.IntegerValue for x in DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElementIds()] or [0])

    try:
        # Esto activa la herramienta de dibujo nativa (con rubber band)
        uidoc.PromptForFamilyInstancePlacement(tipo)
    except OperationCanceledException:
        pass # Es normal, significa que el usuario presionó ESC para terminar

    # Buscamos cuál fue la línea que el usuario acaba de dibujar
    new_elems = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElementIds()
    linea_id = next((eid for eid in new_elems if eid.IntegerValue > max_id_before), None)

    if not linea_id:
        forms.alert("No dibujaste ninguna línea. Script cancelado.", exitscript=True)

    # Extraemos las coordenadas y borramos la línea temporal
    with revit.Transaction("Capturar y limpiar"):
        instancia = doc.GetElement(linea_id)
        curva = instancia.Location.Curve
        p1_capturado = curva.GetEndPoint(0)
        p2_capturado = curva.GetEndPoint(1)
        doc.Delete(linea_id) # Borramos la evidencia
        
    return p1_capturado, p2_capturado

# ===================================================================================
# TRAZADO DE BARRA
# ===================================================================================

# Seleccionar largo de vanos
snaps = UI.Selection.ObjectSnapTypes.Endpoints | UI.Selection.ObjectSnapTypes.Intersections | UI.Selection.ObjectSnapTypes.Nearest # Fuerza la configuración de snaps
# try:
#     p1 = uidoc.Selection.PickPoint(snaps, "Haz clic en el PUNTO INICIAL DEL PRIMER VANO") # Clic 1
#     p2 = uidoc.Selection.PickPoint(snaps, "Haz clic en el PUNTO FINAL DEL PRIMER VANO") # Clic 2
# except OperationCanceledException:
#     forms.alert("Selección de puntos cancelada.", exitscript=True)
p1, p2 = obtener_puntos_con_lapiz("Paso 1: Dibuja la línea del PRIMER VANO")
L1 = ((p2.X-p1.X)**2+(p2.Y-p1.Y)**2)**0.5/3

# try:
#     p1 = uidoc.Selection.PickPoint(snaps, "Haz clic en el PUNTO INICIAL DEL SEGUNDO VANO")
#     p2 = uidoc.Selection.PickPoint(snaps, "Haz clic en el PUNTO FINAL DEL SEGUNDO VANO")
# except OperationCanceledException:
#     forms.alert("Selección de puntos cancelada.", exitscript=True)
p1, p2 = obtener_puntos_con_lapiz("Paso 2: Dibuja la línea del SEGUNDO VANO")
L2 = ((p2.X-p1.X)**2+(p2.Y-p1.Y)**2)**0.5/3

# Seleccionar recorrido
# try:
#     p1 = uidoc.Selection.PickPoint(snaps, "Haz clic en el PUNTO INICIAL DEL RECORRIDO")
#     p2 = uidoc.Selection.PickPoint(snaps, "Haz clic en el PUNTO FINAL DEL RECORRIDO")
# except OperationCanceledException:
#     forms.alert("Selección de puntos cancelada.", exitscript=True)
p1, p2 = obtener_puntos_con_lapiz("Paso 3: Dibuja la línea del RECORRIDO TOTAL")
L_recorrido = ((p2.X-p1.X)**2+(p2.Y-p1.Y)**2)**0.5
ang = math.atan2(p2.Y-p1.Y,p2.X-p1.X)
vdir = (p2 - p1).Normalize() # Vector director normalizado (formato requerido por la API)

# Geometría de la barra
Z = Z_sup - rec_top - UnitUtils.ConvertToInternalUnits(rebar_fi/2, UnitTypeId.Millimeters)
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

# Rebar Set
        N_barras = math.floor(L_recorrido/esp+1)
        rebar.GetShapeDrivenAccessor().SetLayoutAsNumberWithSpacing(
            N_barras,   # Número de barras
            esp,  # Espaciamiento
            True, True, True)   # True = Distribuye hacia la derecha de la línea. Incluir barra inicial. Incluir barra final.
        rebar.SetPresentationMode(uidoc.ActiveView, DB.Structure.RebarPresentationMode.All)

        forms.alert("✅ Rebar Set creado exitosamente. ID: " + rebar.Id.ToString(), title="Éxito")
    except Exception as e:
        print("Error al crear barra: {}".format(e))