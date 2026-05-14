# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from Autodesk.Revit import UI
from Autodesk.Revit.DB import UnitUtils, UnitTypeId
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, forms
doc = revit.doc
uidoc = revit.uidoc

from System import EventHandler, Uri
from System.Windows import SystemParameters
from System.Collections.Generic import List
from System.Windows.Media.Imaging import BitmapImage
import math, os, ctypes
import clr # Permite comunicar Python con C#
clr.AddReference("PresentationFramework")

# ===================================================================================
# PREPARACIÓN DEL ENTORNO
# ===================================================================================
view = uidoc.ActiveView # Vista activa
if view.ViewType != DB.ViewType.EngineeringPlan:
    forms.alert("Esta herramienta solo se puede utilizar en vistas de Planta Estructural.", title="Vista no válida", exitscript=True)

# Crear plano de trabajo
Z_elev = view.GenLevel.Elevation # Obtiene la elevación de la vista
if not view.SketchPlane or abs(abs(view.SketchPlane.GetPlane().Normal.Z) - 1.0) > 0.001 or abs(view.SketchPlane.GetPlane().Origin.Z - Z_elev) > 0.001: # No es necesario crear el plano de trabajo cada vez, si es que ya existe 
    with revit.Transaction("ARAINCO - Configurar Work Plane"):
        plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, Z_elev)) # Se configura un plano horizontal en la elevación Z
        view.SketchPlane = DB.SketchPlane.Create(doc, plane) # Se asigna el plano a la vista activa

# Buscar familia de línea para seleccionar vano
nombre_familia = "EST_D_DEATIL ITEM_DIRECCION VANO MENOR"
tipos = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements()  # Tipos de familia (Family Symbols) cargados en el proyecto
tipo_linea = next((tip for tip in tipos if hasattr(tip, "Family") and tip.Family.Name == nombre_familia), None)  # Se encuentra el tipo que coincide con el nombre indicado
if not tipo_linea:
    forms.alert("No se encontró la familia '{}' en el proyecto.\nPor favor, cárgala antes de usar esta herramienta.".format(nombre_familia), exitscript=True)

# Tipos de barra
rebar_types = DB.FilteredElementCollector(doc).OfClass(DB.Structure.RebarBarType).WhereElementIsElementType().ToElements()  # Lista de barras (Objeto)
rebar_types_names = [rt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for rt in rebar_types]  # Lista de barras (Nombre)

# ===================================================================================
# FUNCIÓN PARA DIBUJAR CON RUBBER BAND
# ===================================================================================
def seleccionar_puntos():
    teclas_enviadas = [False] # Sirve para evitar que se pulse erróneamente varias veces la tecla ESC

    # Se crea una función que captura los datos de la línea que se va a crear
    id_capturado = []
    def capturador(sender, args):
        if args.GetAddedElementIds().Count > 0 and not teclas_enviadas[0]: # Detecta cuando se crea la primera línea
            id_capturado.append(args.GetAddedElementIds()[0])  # Captura el ID de la línea
            teclas_enviadas[0] = True

            # Códigos del sistema operativo para la tecla ESC
            VK_ESCAPE = 0x1B
            KEYEVENTF_KEYUP = 0x0002
            
            # Simulamos presionar y soltar físicamente la tecla ESC (1ra vez)
            ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, 0, 0)
            ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)
            
            # Simulamos presionar y soltar físicamente la tecla ESC (2da vez)
            ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, 0, 0)
            ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)

    # Crear línea
    doc.Application.DocumentChanged += EventHandler[DB.Events.DocumentChangedEventArgs](capturador) # Se activa el capturador en segundo plano
    try:
        uidoc.PromptForFamilyInstancePlacement(tipo_linea)  # Herramienta de dibujo nativa (con rubber band)
    except OperationCanceledException as e: # OperationCanceledException sirve para detectar cuando el usuario cancela la operación de revit (por ej. al apretar ESC)
        error_cancelacion = e 
    finally:
        doc.Application.DocumentChanged -= EventHandler[DB.Events.DocumentChangedEventArgs](capturador)  # Se apaga el capturador

    if not id_capturado:
        raise error_cancelacion or Exception("Operación cancelada") # Detiene la función seleccionar_puntos() y ejecuta el 'except -> continue' desde donde se llamó la función

    linea_id = id_capturado[0]  # Id de la primera línea dibujada
    p1 = doc.GetElement(linea_id).Location.Curve.GetEndPoint(0)
    p2 = doc.GetElement(linea_id).Location.Curve.GetEndPoint(1)
    with revit.Transaction("ARAINCO - Eliminar línea temporal"):
        doc.Delete(linea_id)
    return p1, p2

# ===================================================================================
# FUNCIONES DEL FORMULARIO
# ===================================================================================
class EstadoFormulario: # Almacenar valores del formulario
    def __init__(self):
        self.largo1 = None
        self.largo2 = None
        self.L1 = None
        self.L2 = None
        self.sel_index = -1
        self.esp_text = None
estado = EstadoFormulario()

class ArmaduraSuperiorForm(forms.WPFWindow):  # Funciones del formulario

    # Función al iniciar el formulario
    def __init__(self, xaml_file, datos):
        forms.WPFWindow.__init__(self, xaml_file)
        self.action = None  # Se crea la variable self.action, indicando que aún no se ha seleccionado ningún botón
        self.cmbRebar.ItemsSource = rebar_types_names # Crear ComboBox de lista desplegable con nombres de barra

        # Ruta del logo
        ruta_logo = os.path.join(os.path.dirname(__file__), "logo.png")
        if os.path.exists(ruta_logo):
            self.imgLogo.Source = BitmapImage(Uri(ruta_logo))

        # Posición del formulario en pantalla
        self.Left = 20 # Separación del borde izquierdo
        self.Top = (SystemParameters.WorkArea.Height - self.Height) / 2

        # Restaurar valores seleccionados en el formulario, cada vez que se abra
        if datos.largo1 is not None: 
            self.txtLargo1.Text = "{:.0f}".format(datos.largo1)
        if datos.L1 is not None:     
            self.txtL1.Text = "{:.0f}".format(datos.L1)
        if datos.largo2 is not None: 
            self.txtLargo2.Text = "{:.0f}".format(datos.largo2)
        if datos.L2 is not None:     
            self.txtL2.Text = "{:.0f}".format(datos.L2)
        if 0 <= datos.sel_index < len(rebar_types_names):
            self.cmbRebar.SelectedIndex = datos.sel_index
        if datos.esp_text:              
            self.txtEspaciamiento.Text = datos.esp_text

    # Función para guardar valores de inputs manuales
    def guardar_estado(self):
        estado.sel_index = self.cmbRebar.SelectedIndex
        estado.esp_text = self.txtEspaciamiento.Text

        # Rescatar L1 (si el usuario lo sobreescribió manualmente)
        if self.txtL1.Text:
            try: estado.L1 = float(self.txtL1.Text)
            except ValueError: pass

        # 3. Rescatar L2 (si el usuario lo sobreescribió manualmente)
        if self.txtL2.Text:
            try: estado.L2 = float(self.txtL2.Text)
            except ValueError: pass

    # Funciones al hacer clic en los botones
    def Vano1Click(self, sender, args):
        self.action = "vano1" # Indica que se hizo clic en el botón "vano1"
        self.guardar_estado() # Se guardan datos antes de cerrar el formulario
        self.Close()  # Se cierra el formulario

    def Vano2Click(self, sender, args):
        self.action = "vano2"
        self.guardar_estado()
        self.Close()

    def AplicarClick(self, sender, args):

        # Validación diámetro barra
        if self.cmbRebar.SelectedIndex < 0:
            forms.alert("Debes seleccionar un diámetro de barra.", title="Falta información")
            return
        
        # Validación L1 y L2
        if not self.txtL1.Text or not self.txtL2.Text:
            forms.alert("Debes definir ambos vanos antes de aplicar.", title="Falta información")
            return
        try:
            float(self.txtL1.Text), float(self.txtL2.Text)
        except ValueError:
            forms.alert("Los valores de L1 y L2 deben ser números válidos (solo números, sin letras).", title="Valor inválido")
            return
        
        # Validación espaciamiento
        if not self.txtEspaciamiento.Text:
            forms.alert("Debes ingresar un espaciamiento.", title="Falta información")
            return
        try:
            if float(self.txtEspaciamiento.Text) <= 0:
                forms.alert("El espaciamiento debe ser mayor que cero.", title="Valor inválido")
                return
        except ValueError:
            forms.alert("Ingresa un espaciamiento válido en milímetros.", title="Valor inválido")
            return
        
        self.action = "aplicar"
        self.guardar_estado()
        self.Close()

# ===================================================================================
# FORMULARIO WPF
# ===================================================================================
xaml_path = os.path.join(os.path.dirname(__file__), "ArmaduraSuperiorForm.xaml")

# Bucle de interacción: en cada iteración se crea una nueva ventana, se muestra, y luego se evalúa la acción solicitada.
while True:
    form = ArmaduraSuperiorForm(xaml_path, estado)
    form.ShowDialog() # Se muestra el formulario

    if form.action is None:
        break # Cerrar formulario si es que se apretó la X del formulario o la tecla ESC.

    if form.action == "vano1":
        try: p1, p2 = seleccionar_puntos()
        except OperationCanceledException: continue
        estado.largo1 = UnitUtils.ConvertFromInternalUnits(((p2.X - p1.X) ** 2 + (p2.Y - p1.Y) ** 2) ** 0.5, UnitTypeId.Millimeters)
        estado.L1 = math.ceil(estado.largo1/3/10)*10 # Redondear al centímetro superior
        continue # Permite reiniciar el bucle While True, de modo que se vuelve a abrir el formulario

    if form.action == "vano2":
        try: p1, p2 = seleccionar_puntos()
        except OperationCanceledException: continue
        estado.largo2 = UnitUtils.ConvertFromInternalUnits(((p2.X - p1.X) ** 2 + (p2.Y - p1.Y) ** 2) ** 0.5, UnitTypeId.Millimeters)
        estado.L2 = math.ceil(estado.largo2/3/10)*10
        continue

    if form.action == "aplicar":
        L1 = UnitUtils.ConvertToInternalUnits(estado.L1, UnitTypeId.Millimeters)
        L2 = UnitUtils.ConvertToInternalUnits(estado.L2, UnitTypeId.Millimeters)
        rebar_selected = rebar_types[estado.sel_index]
        rebar_fi = rebar_selected.get_Parameter(DB.BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble()  # Diámetro seleccionado
        esp = UnitUtils.ConvertToInternalUnits(float(estado.esp_text), UnitTypeId.Millimeters)

        try: p1, p2 = seleccionar_puntos()
        except OperationCanceledException: continue

# ===================================================================================
# AUTODETECCIÓN DE LA LOSA O FUNDACIÓN
# ===================================================================================
        
        # Intersección del recorrido con la losa
        filtro_losas = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Floors)
        filtro_fundaciones = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_StructuralFoundation)
        filtro_combinado = DB.LogicalOrFilter(filtro_losas, filtro_fundaciones)
        mid_point = DB.XYZ((p1.X + p2.X) / 2, (p1.Y + p2.Y) / 2, p1.Z) # Punto matemático central del recorrido
        pt_min = DB.XYZ(mid_point.X - 0.0164, mid_point.Y - 0.0164, Z_elev - 4.9213) # Esquina inferior izquierda del pilar
        pt_max = DB.XYZ(mid_point.X + 0.0164, mid_point.Y + 0.0164, Z_elev + 4.9213) # Esquina superior derecha del pilar
        filtro_pilar = DB.BoundingBoxIntersectsFilter(DB.Outline(pt_min, pt_max)) # Pilar rectangular 1x1x300 cm, que que se utiliza para intersectar con la losa
        
        candidatos = DB.FilteredElementCollector(doc, uidoc.ActiveView.Id) \
            .WherePasses(filtro_combinado) \
            .WherePasses(filtro_pilar) \
            .WhereElementIsNotElementType().ToElements()
        if not candidatos:
            forms.alert("No se encontró ninguna losa visible en un rango de ±1.5 metros, respecto al nivel de la vista actual.", title="Losa no encontrada")
            continue
        
        # Plan de Rescate: Asumimos la primera losa candidata y su cota superior, en caso de que falle la proyección
        slab = candidatos[0]
        Z_sup = slab.get_BoundingBox(uidoc.ActiveView).Max.Z
        plane_normal = DB.XYZ.BasisZ  # Normal apuntando hacia arriba
        plane_origin = DB.XYZ(0, 0, Z_sup) # Origen en la altura máxima

        # Analizar geometría de cada losa candidata
        opt = DB.Options()
        opt.DetailLevel = DB.ViewDetailLevel.Coarse # Usa la geometría más simple y rápida posible
        p_test = DB.XYZ(mid_point.X, mid_point.Y, pt_max.Z) # Punto usado para generar láser de proyección hacia la cara de la losa candidata
        dist_min = float('inf')

        for candidato in candidatos:
            if not candidato.get_Geometry(opt): continue # Si el elemento está corrupto o no tiene masa 3D, se avanza al siguiente candidato

            cara_encontrada = False
            for geom_obj in candidato.get_Geometry(opt): # Itera los bloques de geometría (info de sólidos, líneas y puntos de la losa)
                if cara_encontrada: break # Si ya hallamos la cara, aborta la revisión de este elemento

                if isinstance(geom_obj, DB.GeometryInstance):
                    iterable_geom = geom_obj.GetInstanceGeometry() # Si es una familia (ej. zapata aislada) o un model in-place, Revit lo trata como un bloque o símbolo
                else: 
                    iterable_geom = [geom_obj] # Si es una losa nativa de Revit

                for solid in iterable_geom:
                    if cara_encontrada: break # Si ya hallamos la cara, aborta la revisión de los demás sólidos

                    if isinstance(solid, DB.Solid) and solid.Faces.Size > 0: # Filtra solo las piezas sólidas (descarta líneas y puntos)
                        for face in solid.Faces: # Itera las caras individuales de las piezas sólidas (6 caras)
                            if isinstance(face, DB.PlanarFace) and face.FaceNormal.Z > 0.5: # Busca solo caras planas que apunten hacia arriba (tolerando rampas/pendientes)
                                proj = face.Project(p_test) # Intersección del rayo vertical en la cara superior de la losa
                                if proj: # Revisa si la proyección es válida. Si falla, el código simplemente usará la elevación del "Plan de Rescate". Para evitar que falle, las losas deben tener prioridad de unión por sobre las vigas, columnas y muros
                                    dist = abs(proj.XYZPoint.Z - Z_elev)
                                    if dist < dist_min:
                                        dist_min = dist
                                        slab = candidato
                                        plane_normal = face.FaceNormal
                                        plane_origin = face.Origin
                                        cara_encontrada = True
                                        break
        
        # Parámetros de la losa encontrada
        try:
            rec_top_param = slab.get_Parameter(DB.BuiltInParameter.CLEAR_COVER_TOP)
            rec_top = doc.GetElement(rec_top_param.AsElementId()).CoverDistance
        except Exception:
            forms.alert("La losa detectada no tiene parámetros de recubrimiento estructural válidos.", title="Error de Parámetro")
            continue
        
# ===================================================================================
# GEOMETRÍA DE LA BARRA
# ===================================================================================
        Z = Z_sup - rec_top - rebar_fi / 2
        ang = math.atan2(p2.Y - p1.Y, p2.X - p1.X)
        p1 = DB.XYZ(p1.X + (rec_top + rebar_fi / 2) * math.cos(ang), p1.Y + (rec_top + rebar_fi / 2) * math.sin(ang), Z) # Dejar la barra con recubrimiento lateral al borde de la losa
        p2 = DB.XYZ(p2.X - (rec_top + rebar_fi / 2) * math.cos(ang), p2.Y - (rec_top + rebar_fi / 2) * math.sin(ang), Z)
        L_recorrido = ((p2.X - p1.X) ** 2 + (p2.Y - p1.Y) ** 2) ** 0.5
        vdir = (p2 - p1).Normalize() # Vector director normalizado (formato requerido por la API)

        p_start = DB.XYZ(p1.X - L1 * math.sin(ang), p1.Y + L1 * math.cos(ang), Z) # Punto inicial de la barra
        p_end = DB.XYZ(p1.X + L2 * math.sin(ang), p1.Y - L2 * math.cos(ang), Z) # Punto final de la barra
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
                    rebar_selected, # Tipo de barra
                    None, None, # Gancho inicial y final
                    slab, # Anfitrión
                    vdir, # Vector director
                    curves,
                    DB.Structure.RebarHookOrientation.Right, # Orientación de gancho inicial
                    DB.Structure.RebarHookOrientation.Right, # Orientación de gancho final
                    True, True) # Usar Rebar Shape existente. Si no existe Rebar Shape que coincida, Revit creará un Rebar Shape personalizado
                
        # Transformar a Rebar Set
                rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(
                    esp,
                    L_recorrido,
                    True, True, True) # True = Distribuye hacia la derecha de la línea. Incluir barra inicial. Incluir barra final.
                doc.Regenerate()

                # Espaciamiento real con el que quedan las barras
                subelementos = list(rebar.GetSubelements()) # Se extrae cada barra individual del Rebar Set
                if len(subelementos) > 2:
                    esp_real = L_recorrido / (len(subelementos) - 1)
                    rebar.SetPresentationMode(uidoc.ActiveView, DB.Structure.RebarPresentationMode.Middle)
                else:
                    esp_real = 0 # Por si es que en el recorrido indicado solo caben 1 o 2 barras
                    rebar.SetPresentationMode(uidoc.ActiveView, DB.Structure.RebarPresentationMode.All)

        # Offset para Multi-Rebar Annotation y Tag
                if L1>L2:
                    offset_mra = UnitUtils.ConvertToInternalUnits(-80, UnitTypeId.Centimeters)
                    offset_tag = UnitUtils.ConvertToInternalUnits(175, UnitTypeId.Centimeters)
                else:
                    offset_mra = UnitUtils.ConvertToInternalUnits(80, UnitTypeId.Centimeters)
                    offset_tag = UnitUtils.ConvertToInternalUnits(-175, UnitTypeId.Centimeters)

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
                    mid_point = DB.XYZ((p1.X + p2.X) / 2 - offset_mra * math.sin(ang), (p1.Y + p2.Y) / 2 + offset_mra * math.cos(ang), Z) # Se centra la cota al punto medio del recorrido y luego se suma el offset
                    tipo_opts.DimensionLineOrigin = mid_point
                    tipo_opts.TagHeadPosition = mid_point
                    DB.MultiReferenceAnnotation.Create(doc, uidoc.ActiveView.Id, tipo_opts)
                else:
                    forms.alert("La armadura se creó, pero no se encontró el tipo de cota llamado '{}'.\n\nRevisa que esté cargado en el proyecto o que el nombre esté escrito exactamente igual.".format(nombre_tipo), title="Falta Tipo de Recorrido")

        # Tag
                nombre_tag = "Marca - Cantidad - Diametro - Espaciamiento"
                tipos = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_RebarTags) # Tipos de etiqueta en el proyecto
                tipo = next((t for t in tipos if DB.Element.Name.GetValue(t) == nombre_tag), None) # Se encuentra el tipo que coincide con el nombre indicado

                # Encontrar la barra central para etiquetar
                idx_central = (len(subelementos)-1) // 2
                rebar_central = subelementos[idx_central] # Se obtiene la barra central
                centro_primera = DB.XYZ((p_start.X + p_end.X) / 2, (p_start.Y + p_end.Y) / 2, Z) # Centroide de la primera barra
                centro_barra = DB.XYZ(centro_primera.X + vdir.X * idx_central * esp_real - offset_tag * math.sin(ang), centro_primera.Y + vdir.Y * idx_central * esp_real + offset_tag * math.cos(ang), Z) # Centroide de la barra central

                if tipo:
                    tag = DB.IndependentTag.Create(
                        doc,
                        tipo.Id,
                        uidoc.ActiveView.Id,
                        rebar_central.GetReference(),
                        False, # ¿Llevará directriz/Leader Line?
                        DB.TagOrientation.AnyModelDirection,
                        centro_barra) # Posición
                    tag.RotationAngle = ang + (math.pi / 2)
                else:
                    forms.alert("La armadura se creó, pero no se encontró el tipo de etiqueta llamado '{}'.\n\nRevisa que esté cargado en el proyecto o que el nombre esté escrito exactamente igual.".format(nombre_tag), title="Falta Etiqueta")

         # Parámetros
                param = "Armadura_Ubicacion"
                if rebar.LookupParameter(param) and not rebar.LookupParameter(param).IsReadOnly:
                    rebar.LookupParameter(param).Set("F's")
                else:
                    forms.alert("No se encontró el parámetro de instancia '{}', o está bloqueado.".format(param), title="Falta parámetro")

            except Exception as e:
                print("Error al crear la barra: {}".format(e))
        continue # Permite aplicar varios recorridos, al reiniciar el bucle While True
