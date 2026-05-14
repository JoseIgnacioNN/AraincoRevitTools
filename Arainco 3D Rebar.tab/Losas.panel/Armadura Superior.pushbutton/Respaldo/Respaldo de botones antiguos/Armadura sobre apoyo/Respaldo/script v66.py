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

# Definición del nivel
view = uidoc.ActiveView # Vista activa
try:
    view_range = view.GetViewRange()
    top_level_id = view_range.GetLevelId(DB.PlanViewPlane.TopClipPlane)
    if top_level_id != DB.ElementId.InvalidElementId:
        Z_elev = doc.GetElement(top_level_id).Elevation
    else:
        Z_elev = view.GenLevel.Elevation # Respaldo si el "Top" está como "Unlimited"
except Exception:
    Z_elev = view.GenLevel.Elevation # Respaldo si Revit arroja error

# Crear plano de trabajo
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
rebar_types = sorted(rebar_types, key=lambda rt: rt.get_Parameter(DB.BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble()) # Ordenar la lista por diámetro de barra
rebar_types_names = [rt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for rt in rebar_types]  # Lista de barras (Nombre)

# Función auxiliar para resolver la ecuación del plano
def get_z_on_plane(x, y, origen, normal):
    return origen.Z - (normal.X * (x - origen.X) + normal.Y * (y - origen.Y)) / normal.Z

# ===================================================================================
# FUNCIÓN PARA DIBUJAR CON RUBBER BAND
# ===================================================================================
def seleccionar_puntos(cant_lineas):
    teclas_enviadas = [False] # Sirve para evitar que se pulse erróneamente varias veces la tecla ESC

    # Se crea una función que captura los datos de la línea que se va a crear
    id_capturado = []
    def capturador(sender, args):
        for elem_id in args.GetAddedElementIds(): # Recorremos todas las líneas creadas
            if elem_id not in id_capturado:
                id_capturado.append(elem_id) # Captura el ID de la línea

        # Simular apretar la tecla ESC 2 veces, cuando ya tenemos la cantidad de líneas solicitadas
        if len(id_capturado) >= cant_lineas and not teclas_enviadas[0]: 
            teclas_enviadas[0] = True
            VK_ESCAPE = 0x1B
            KEYEVENTF_KEYUP = 0x0002
            ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, 0, 0)
            ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)
            ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, 0, 0)
            ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)

    # Transaction Group "Fantasma". Oculta la transacción en el listado Ctrl+Z de Revit
    t_ghost = DB.TransactionGroup(doc, "Borrador de Rastro")
    t_ghost.Start()

    # Crear línea
    doc.Application.DocumentChanged += EventHandler[DB.Events.DocumentChangedEventArgs](capturador) # Se activa el capturador en segundo plano
    try:
        uidoc.PromptForFamilyInstancePlacement(tipo_linea)  # Herramienta de dibujo nativa (con rubber band)
    except OperationCanceledException as e: # OperationCanceledException sirve para detectar cuando el usuario cancela la operación de revit (por ej. al apretar ESC)
        error_cancelacion = e 
    finally:
        doc.Application.DocumentChanged -= EventHandler[DB.Events.DocumentChangedEventArgs](capturador)  # Se apaga el capturador

    # Guardar estado de pantalla
    active_uiview = next((u for u in uidoc.GetOpenUIViews() if u.ViewId == uidoc.ActiveView.Id), None)
    zoom_corners = active_uiview.GetZoomCorners() if active_uiview else None

    # Si el usuario no dibujó la cantidad requerida de líneas (apretó ESC antes)
    if len(id_capturado) < cant_lineas:
        t_ghost.RollBack() # Borra la línea sin dejar rastro en el historial
        if zoom_corners: active_uiview.ZoomAndCenterRectangle(zoom_corners[0], zoom_corners[1]) # Recuperar estado de pantalla
        raise error_cancelacion or Exception("Operación cancelada") # Detiene la función seleccionar_puntos() y ejecuta el 'except -> continue' desde donde se llamó la función

    # Extracción de los puntos de la línea 1
    p1 = doc.GetElement(id_capturado[0]).Location.Curve.GetEndPoint(0)
    p2 = doc.GetElement(id_capturado[0]).Location.Curve.GetEndPoint(1)
    p1 = DB.XYZ(p1.X, p1.Y, Z_elev) # Se fuerza la coordenada Z = Z_elev, para evitar un snap accidental a una altura distinta
    p2 = DB.XYZ(p2.X, p2.Y, Z_elev)

    # Extracción de los puntos de la línea 2 (solo si se pidió)
    if cant_lineas == 2:
        p3 = doc.GetElement(id_capturado[1]).Location.Curve.GetEndPoint(0)
        p4 = doc.GetElement(id_capturado[1]).Location.Curve.GetEndPoint(1)
        p3 = DB.XYZ(p3.X, p3.Y, Z_elev)
        p4 = DB.XYZ(p4.X, p4.Y, Z_elev)

    t_ghost.RollBack()
    if zoom_corners: active_uiview.ZoomAndCenterRectangle(zoom_corners[0], zoom_corners[1])
    if cant_lineas == 2:
        return p1, p2, p3, p4
    else:
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
        self.Lext = None
        self.Lext_izq = True
        self.sel_index = -1
        self.esp_text = None
estado = EstadoFormulario()

class Formulario(forms.WPFWindow):  # Funciones del formulario

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
        self.Left = 5 # Separación del borde izquierdo
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
        if datos.Lext is not None:
            self.txtLext.Text = "{:.0f}".format(datos.Lext)
        if datos.Lext_izq is not None:
            self.chkIzq.IsChecked = datos.Lext_izq
            self.chkDer.IsChecked = not datos.Lext_izq
        if 0 <= datos.sel_index < len(rebar_types_names):
            self.cmbRebar.SelectedIndex = datos.sel_index
        if datos.esp_text:              
            self.txtEspaciamiento.Text = datos.esp_text

    # Función para guardar valores de inputs manuales
    def guardar_estado(self):
        estado.sel_index = self.cmbRebar.SelectedIndex
        estado.esp_text = self.txtEspaciamiento.Text
        estado.Lext_izq = self.chkIzq.IsChecked

        # Rescatar L1 (si el usuario lo sobreescribió manualmente)
        if self.txtL1.Text:
            try: estado.L1 = float(self.txtL1.Text)
            except ValueError: pass

        # Rescatar L2 (si el usuario lo sobreescribió manualmente)
        if self.txtL2.Text:
            try: estado.L2 = float(self.txtL2.Text)
            except ValueError: pass

        # Rescatar L extensión (si el usuario lo sobreescribió manualmente)
        if self.txtLext.Text:
            try: estado.Lext = float(self.txtLext.Text)
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

    def LextClick(self, sender, args):
        self.action = "Lext"
        self.guardar_estado()
        self.Close()

    def ChkIzq_Click(self, sender, args):
        self.chkIzq.IsChecked = True
        self.chkDer.IsChecked = False

    def ChkDer_Click(self, sender, args):
        self.chkIzq.IsChecked = False
        self.chkDer.IsChecked = True

    def AplicarClick(self, sender, args):

        # Validación diámetro barra
        if self.cmbRebar.SelectedIndex < 0:
            forms.alert("Debes seleccionar un diámetro de barra.", title="Falta información")
            return
        
        # Validación L1 y L2
        if not self.txtL1.Text or not self.txtL2.Text:
            forms.alert("Debes definir los largos de barra antes de aplicar.", title="Falta información")
            return
        try:
            if float(self.txtL1.Text) <= 0 or float(self.txtL2.Text) <= 0:
                forms.alert("Los largos de barra deben ser mayor que cero.", title="Valor inválido")
                return
        except ValueError:
            forms.alert("Los largos de barra deben ser números válidos.", title="Valor inválido")
            return
        
        # Validación L extensión
        if self.txtLext.Text:
            try: 
                if float(self.txtLext.Text) < 0:
                    forms.alert("La extensión de la armadura debe ser mayor que cero.", title="Valor inválido")
                    return
            except ValueError:
                forms.alert("El largo de extensión debe ser un número válido.", title="Valor inválido")
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
            forms.alert("Ingresa un espaciamiento válido.", title="Valor inválido")
            return
        
        self.action = "aplicar"
        self.guardar_estado()
        self.Close()

# ===================================================================================
# FORMULARIO WPF
# ===================================================================================
xaml_path = os.path.join(os.path.dirname(__file__), "Formulario.xaml")

# Bucle de interacción: en cada iteración se crea una nueva ventana, se muestra, y luego se evalúa la acción solicitada.
while True:
    form = Formulario(xaml_path, estado)
    form.ShowDialog() # Se muestra el formulario

    if form.action is None:
        break # Cerrar formulario si es que se apretó la X del formulario o la tecla ESC.

    if form.action == "vano1":
        try: p1, p2 = seleccionar_puntos(1)
        except OperationCanceledException: continue
        estado.largo1 = UnitUtils.ConvertFromInternalUnits((p2 - p1).GetLength(), UnitTypeId.Millimeters)
        estado.L1 = math.ceil(estado.largo1/3/10)*10 # Redondear al centímetro superior
        continue # Permite reiniciar el bucle While True, de modo que se vuelve a abrir el formulario

    if form.action == "vano2":
        try: p1, p2 = seleccionar_puntos(1)
        except OperationCanceledException: continue
        estado.largo2 = UnitUtils.ConvertFromInternalUnits((p2 - p1).GetLength(), UnitTypeId.Millimeters)
        estado.L2 = math.ceil(estado.largo2/3/10)*10
        continue

    if form.action == "Lext":
        try: p1, p2 = seleccionar_puntos(1)
        except OperationCanceledException: continue
        estado.Lext = round(UnitUtils.ConvertFromInternalUnits((p2 - p1).GetLength(), UnitTypeId.Millimeters)/10)*10
        continue

    if form.action == "aplicar":
        L1 = UnitUtils.ConvertToInternalUnits(estado.L1, UnitTypeId.Millimeters)
        L2 = UnitUtils.ConvertToInternalUnits(estado.L2, UnitTypeId.Millimeters)
        Lext = UnitUtils.ConvertToInternalUnits(estado.Lext if estado.Lext else 0, UnitTypeId.Millimeters)
        Lext_izq = estado.Lext_izq
        rebar_selected = rebar_types[estado.sel_index]
        rebar_fi = rebar_selected.get_Parameter(DB.BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble()  # Diámetro seleccionado
        esp = UnitUtils.ConvertToInternalUnits(float(estado.esp_text), UnitTypeId.Millimeters)

        try: p1, p2 = seleccionar_puntos(1)
        except OperationCanceledException: continue

        # Transaction Group
        t_group = DB.TransactionGroup(doc, "ARAINCO - Armadura Superior sobre apoyo")
        t_group.Start()

# ===================================================================================
# AUTODETECCIÓN DE LA LOSA O FUNDACIÓN
# ===================================================================================
        
        # Intersección del recorrido con la losa
        filtro_combinado = DB.LogicalOrFilter(DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Floors), 
                                              DB.ElementCategoryFilter(DB.BuiltInCategory.OST_StructuralFoundation))
        mid_point = (p1 + p2) / 2 # Punto matemático central del recorrido
        pt_min = DB.XYZ(mid_point.X - 0.0164, mid_point.Y - 0.0164, Z_elev - 6.562) # Esquina inferior izquierda del pilar
        pt_max = DB.XYZ(mid_point.X + 0.0164, mid_point.Y + 0.0164, Z_elev + 6.562) # Esquina superior derecha del pilar
        filtro_pilar = DB.BoundingBoxIntersectsFilter(DB.Outline(pt_min, pt_max)) # Pilar rectangular 1x1x400 cm, que que se utiliza para intersectar con la losa
        
        # Intersección inicial y rápida
        candidatos = DB.FilteredElementCollector(doc, uidoc.ActiveView.Id) \
            .WherePasses(filtro_combinado) \
            .WherePasses(filtro_pilar) \
            .WhereElementIsNotElementType().ToElements()
        if not candidatos:
            t_group.RollBack()
            forms.alert("No se encontró ninguna losa visible en un rango de ±2 metros respecto al nivel de la vista actual. Armadura no creada.", title="Losa no encontrada")
            continue

        # Intersección final y detallada. Analiza la geometría de cada losa candidata
        opt = DB.Options()
        opt.DetailLevel = DB.ViewDetailLevel.Coarse # Usa la geometría más simple y rápida posible
        p_test = DB.XYZ(mid_point.X, mid_point.Y, Z_elev) # Punto usado para generar láser de proyección hacia la cara de la losa candidata
        dist_min = float('inf')
        slab = None
        host_validos = set() # Servirá para contar los host validos

        for candidato in candidatos:
            if not candidato.get_Geometry(opt): continue # Si el elemento está corrupto o no tiene masa 3D, se avanza al siguiente candidato

            for geom_obj in candidato.get_Geometry(opt): # Itera los bloques de geometría (info de sólidos, líneas y puntos de la losa)
                if isinstance(geom_obj, DB.GeometryInstance):
                    iterable_geom = geom_obj.GetInstanceGeometry() # Si es una familia (ej. zapata aislada) o un model in-place, Revit lo trata como un bloque o símbolo
                else: 
                    iterable_geom = [geom_obj] # Si es una losa nativa de Revit

                for solid in iterable_geom:
                    if isinstance(solid, DB.Solid) and solid.Faces.Size > 0: # Filtra solo las piezas sólidas (descarta líneas y puntos)

                        for face in solid.Faces: # Itera las caras individuales de las piezas sólidas (6 caras)
                            if isinstance(face, DB.PlanarFace) and face.FaceNormal.Z > 0.5: # Busca solo caras planas que apunten hacia arriba (tolerando rampas/pendientes)
                                
                                # Si la proyección devuelve None, el punto cae sobre un shaft o sobre una viga, muro o columna que tenga prioridad por sobre la losa
                                interseccion_fisica = face.Project(p_test)
                                if not interseccion_fisica: 
                                    continue
                                
                                host_validos.add(candidato.Id)
                                Z_losa = get_z_on_plane(mid_point.X, mid_point.Y, face.Origin, face.FaceNormal) # Elevación de la losa en la intersección con el pilar
                                dist = abs(Z_losa - Z_elev)
                                if dist < dist_min:
                                    dist_min = dist
                                    slab = candidato
                                    plane_normal = face.FaceNormal
                                    plane_origin = face.Origin
                                    break
        
        if len(host_validos) > 1:
            forms.alert("Se detectó más de una posible losa anfitrión. Se asignó la losa más cercana a la elevación de la planta.", title="Losa anfitrión")
        
        if not slab:
            t_group.RollBack()
            mensaje = """No se pudo detectar un host válido, por alguna de las siguientes razones:
• La armadura se encuentra sobre un shaft.
• La losa no tiene prioridad de unión por sobre la viga, columna o muro.
• La losa no tiene caras superiores válidas."""
            forms.alert(mensaje, title="Host no encontrado")
            continue
        
        # Parámetros de la losa encontrada
        try:
            rec_top_param = slab.get_Parameter(DB.BuiltInParameter.CLEAR_COVER_TOP)
            rec_top = doc.GetElement(rec_top_param.AsElementId()).CoverDistance
        except Exception:
            t_group.RollBack()
            forms.alert("La losa detectada no tiene parámetros de recubrimiento estructural válidos.", title="Error de Parámetro")
            continue
        
# ===================================================================================
# GEOMETRÍA DE LA BARRA
# ===================================================================================
        offset_lateral = rec_top + rebar_fi / 2
        offset_rec = rec_top + rebar_fi / 2
        origen_desplazado = plane_origin - plane_normal * offset_rec # Desplazamos el plano matemático hacia abajo para incluir el recubrimiento y diámetro
        
        # Evitar que el código colapse si el usuario dibuja una línea muy corta
        if (p2-p1).GetLength() <= (2 * offset_lateral + rebar_fi):
            t_group.RollBack()
            forms.alert("El recorrido dibujado es demasiado corto para aplicar los recubrimientos laterales.", title="Recorrido muy corto")
            continue 
        
        v_recorrido_2D = (p2-p1).Normalize()
        p1 = p1 + v_recorrido_2D * offset_lateral # Punto inicial del recorrido, reducido por el recubrimiento lateral del borde de la losa
        p2 = p2 - v_recorrido_2D * offset_lateral
        p1_z = get_z_on_plane(p1.X, p1.Y, origen_desplazado, plane_normal)
        p2_z = get_z_on_plane(p2.X, p2.Y, origen_desplazado, plane_normal)
        p1_3D = DB.XYZ(p1.X, p1.Y, p1_z)
        p2_3D = DB.XYZ(p2.X, p2.Y, p2_z)
        v_recorrido_3D = (p2_3D - p1_3D).Normalize() # Vector director del recorrido en 3D. Es perpendicular al eje de la barra y paralelo a la pendiente de la losa
        v_bar_3D = plane_normal.CrossProduct(v_recorrido_3D).Normalize() # # Vector director de la barra (en la dirección longitudinal)

        if Lext_izq == True:
            start = p1_3D + v_bar_3D * (L1 + Lext) # Extremo inicial de la barra
            end = p1_3D - v_bar_3D * L2 # Extremo final de la barra
        else:
            start = p1_3D + v_bar_3D * L1
            end = p1_3D - v_bar_3D * (L2 + Lext)

        curves = [DB.Line.CreateBound(start, end)] # Las puntos que conforman la barra deben estar en una lista
        L_barra = (end - start).GetLength()
        L_recorrido = (p2_3D - p1_3D).GetLength()

# ===================================================================================
# CREAR LA BARRA
# ===================================================================================
        t = DB.Transaction(doc, "ARAINCO - Crear Rebar Set")
        t.Start()
        try:

    # Barra individual
            rebar = DB.Structure.Rebar.CreateFromCurves(
                doc,
                DB.Structure.RebarStyle.Standard, # Estilo de barra (Standard o StirrupTie)
                rebar_selected, # Tipo de barra
                None, None, # Gancho inicial y final
                slab, # Anfitrión
                v_recorrido_3D, # Vector director
                curves,
                DB.Structure.RebarHookOrientation.Left, # Orientación de gancho inicial
                DB.Structure.RebarHookOrientation.Left, # Orientación de gancho final
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
            if Lext == 0:
                if L1 > L2:
                    offset_mra = UnitUtils.ConvertToInternalUnits(-65, UnitTypeId.Centimeters)
                    offset_tag = UnitUtils.ConvertToInternalUnits(175, UnitTypeId.Centimeters)
                else:
                    offset_mra = UnitUtils.ConvertToInternalUnits(65, UnitTypeId.Centimeters)
                    offset_tag = UnitUtils.ConvertToInternalUnits(-175, UnitTypeId.Centimeters)
            elif Lext < UnitUtils.ConvertToInternalUnits(290, UnitTypeId.Centimeters):
                if L1 > L2:
                    if Lext_izq == True:
                        offset_mra = UnitUtils.ConvertToInternalUnits(65, UnitTypeId.Centimeters) + Lext
                        offset_tag = UnitUtils.ConvertToInternalUnits(175, UnitTypeId.Centimeters)
                    else:
                        offset_mra = UnitUtils.ConvertToInternalUnits(65, UnitTypeId.Centimeters)
                        offset_tag = UnitUtils.ConvertToInternalUnits(175, UnitTypeId.Centimeters)
                else:
                    if Lext_izq == True:
                        offset_mra = UnitUtils.ConvertToInternalUnits(-65, UnitTypeId.Centimeters)
                        offset_tag = UnitUtils.ConvertToInternalUnits(-175, UnitTypeId.Centimeters)
                    else:
                        offset_mra = UnitUtils.ConvertToInternalUnits(-65, UnitTypeId.Centimeters) - Lext
                        offset_tag = UnitUtils.ConvertToInternalUnits(-175, UnitTypeId.Centimeters)
            else:
                if Lext_izq == True:
                    offset_mra = UnitUtils.ConvertToInternalUnits(65, UnitTypeId.Centimeters)
                    offset_tag = UnitUtils.ConvertToInternalUnits(160, UnitTypeId.Centimeters) - L_barra/2 + L2
                else:
                    offset_mra = UnitUtils.ConvertToInternalUnits(-65, UnitTypeId.Centimeters)
                    offset_tag = UnitUtils.ConvertToInternalUnits(-160, UnitTypeId.Centimeters) + L_barra/2 - L1

    # Multi-Rebar Annotation
            nombre_tipo = "Recorrido Barras"
            tipos = DB.FilteredElementCollector(doc).OfClass(DB.MultiReferenceAnnotationType)
            tipo = next((t for t in tipos if DB.Element.Name.GetValue(t) == nombre_tipo), None) # Se encuentra la cota indicada

            if tipo:
                tipo_opts = DB.MultiReferenceAnnotationOptions(tipo) # Se abre la configuración de Multi-Rebar Annotation
                rebar_ids = List[DB.ElementId]()
                rebar_ids.Add(rebar.Id)
                tipo_opts.SetElementsToDimension(rebar_ids)

                # El vector director de la cota debe ser 2D
                v_bar_2D = DB.XYZ(v_bar_3D.X, v_bar_3D.Y, 0).Normalize()
                v_mra_2D = DB.XYZ(-v_bar_2D.Y, v_bar_2D.X, 0)
                if v_mra_2D.DotProduct(v_recorrido_2D) < 0: # Nos aseguramos de que no apunte al lado contrario
                    v_mra_2D = -v_mra_2D

                tipo_opts.DimensionLineDirection = v_mra_2D
                tipo_opts.DimensionPlaneNormal = uidoc.ActiveView.ViewDirection
                mid_point = (p1 + p2) / 2 + v_bar_2D * offset_mra # Se centra la cota al punto medio del recorrido y luego se suma el offset
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
            centro_primera = centro_primera = (start + end) / 2 # Centroide de la primera barra
            centro_barra = centro_primera + v_recorrido_3D * idx_central * esp_real # Centroide de la barra central
            tag_pos = DB.XYZ(centro_barra.X, centro_barra.Y, Z_elev) + v_bar_2D * offset_tag

            if tipo:
                try:
                    tag = DB.IndependentTag.Create(
                        doc,
                        tipo.Id,
                        uidoc.ActiveView.Id,
                        rebar_central.GetReference(),
                        False, # ¿Llevará directriz/Leader Line?
                        DB.TagOrientation.AnyModelDirection, # La familia del tag debe tener activada la casilla "Rotate with component". De lo contrario, hay que asignar la rotación manualmente
                        tag_pos)
                except Exception:
                    forms.alert("Se creó la armadura, pero no fue posible etiquetarla", title="Error de etiqueta")
            else:
                forms.alert("La armadura se creó, pero no se encontró el tipo de etiqueta llamado '{}'.\n\nRevisa que esté cargado en el proyecto o que el nombre esté escrito exactamente igual.".format(nombre_tag), title="Falta Etiqueta")

        # Parámetros y Visibilidad
            param = "Armadura_Ubicacion"
            if rebar.LookupParameter(param) and not rebar.LookupParameter(param).IsReadOnly:
                rebar.LookupParameter(param).Set("F's")
            else:
                forms.alert("No se encontró el parámetro de instancia '{}', o está bloqueado.".format(param), title="Error de parámetro")

            param = "Armadura_Arainco"
            if rebar.LookupParameter(param) and not rebar.LookupParameter(param).IsReadOnly:
                rebar.LookupParameter(param).Set(1)
            else:
                forms.alert("No se encontró el parámetro de instancia '{}', o está bloqueado.".format(param), title="Error de parámetro")

            rebar.SetUnobscuredInView(view, True)

            t.Commit()
            t_group.Assimilate()
            continue # Permite aplicar varios recorridos, al reiniciar el bucle While True

        except Exception as e:
            if t.HasStarted(): t.RollBack()
            if t_group.HasStarted(): t_group.RollBack()
            forms.alert("Error al crear la barra: {}".format(e), title="Error de creación de barra")
            continue
        
