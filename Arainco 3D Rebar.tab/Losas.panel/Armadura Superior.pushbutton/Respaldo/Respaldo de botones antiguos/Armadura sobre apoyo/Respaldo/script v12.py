# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from Autodesk.Revit import UI
from Autodesk.Revit.DB import UnitUtils, UnitTypeId
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, forms
from System.Collections.Generic import List
import math
import os
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

bbox = slab.get_BoundingBox(None) # Caja geométrica que encierra toda la losa
Z_sup = bbox.Max.Z # Elevación de cara superior
e = slab.get_Parameter(DB.BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble() # Espesor
rec_top_param = slab.get_Parameter(DB.BuiltInParameter.CLEAR_COVER_TOP) # Parámetro del recubrimiento superior
rec_top = doc.GetElement(rec_top_param.AsElementId()).CoverDistance # Recubrimiento superior

# ===================================================================================
# PREPARACIÓN DEL ENTORNO
# ===================================================================================
# Crear plano de trabajo
view = uidoc.ActiveView  # Vista activa
if not view.SketchPlane or abs(view.SketchPlane.GetPlane().Normal.Z) != 1:  # No es necesario crear el plano de trabajo cada vez, si es que ya existe
    with revit.Transaction("ARAINCO - Configurar Work Plane"):
        plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ.Zero)  # Se configura un plano horizontal en el origen (Z=0)
        view.SketchPlane = DB.SketchPlane.Create(doc, plane)  # Se asigna el plano a la vista activa

# Buscar familia de línea para seleccionar vano
nombre_familia = "EST_D_DEATIL ITEM_EMPALME"
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
    id_capturado = []

    # Se crea una función que captura los datos de la línea que se va a crear
    def capturador(sender, args):
        nuevos_ids = args.GetAddedElementIds()
        if nuevos_ids.Count > 0:
            id_capturado.append(nuevos_ids[0])  # Captura el ID de la línea
    doc.Application.DocumentChanged += capturador  # Se activa la función en segundo plano

    # Crear línea (instancia de familia de detalle de línea)
    try:
        uidoc.PromptForFamilyInstancePlacement(tipo_linea)  # Herramienta de dibujo nativa (con rubber band)
    except OperationCanceledException:
        pass  # Avanzar con el código si el usuario presionó ESC para terminar
    finally:
        doc.Application.DocumentChanged -= capturador  # Se apaga la función en segundo plano

    if not id_capturado:
        forms.alert("Operación cancelada. No dibujaste la línea.", exitscript=True)

    # Se extraen las coordenadas y se borra la línea temporal
    linea_id = id_capturado[0]  # Id de la primera línea dibujada
    p1 = doc.GetElement(id_capturado[0]).Location.Curve.GetEndPoint(0)
    p2 = doc.GetElement(id_capturado[0]).Location.Curve.GetEndPoint(1)
    with revit.Transaction("ARAINCO - Selección de puntos"):
        doc.Delete(linea_id)  # Se borra la línea temporal

    return p1, p2

# ===================================================================================
# FORMULARIO WPF
# ===================================================================================
class ArmaduraSuperiorForm(forms.WPFWindow):
    def __init__(self, xaml_file, L1_v1=None, L1_v2=None, largo1_cm=None, largo2_cm=None, L1_cm=None, L2_cm=None, sel_index=-1, esp_text=None):
        forms.WPFWindow.__init__(self, xaml_file)

        # Guardar acción a ejecutar fuera del formulario
        self.action = None  # "vano1", "vano2", "aplicar", "cancelar"

        # Cargar lista de tipos de barra
        self.cmbRebar.ItemsSource = rebar_types_names
        if sel_index >= 0 and sel_index < len(rebar_types_names):
            self.cmbRebar.SelectedIndex = sel_index

        # Restaurar longitudes (si ya existen)
        if largo1_cm is not None:
            self.txtLargo1.Text = "{:.1f}".format(largo1_cm)
        if L1_cm is not None:
            self.txtL11.Text = "{:.1f}".format(L1_cm)
        if largo2_cm is not None:
            self.txtLargo2.Text = "{:.1f}".format(largo2_cm)
        if L2_cm is not None:
            self.txtL12.Text = "{:.1f}".format(L2_cm)

        # Restaurar espaciamiento si ya existe (en texto)
        if esp_text:
            self.txtEspaciamiento.Text = esp_text

        # Valores internos en unidades de Revit
        self.L1_v1 = L1_v1
        self.L1_v2 = L1_v2

    def Vano1Click(self, sender, args):
        self.action = "vano1"
        self.Close()  # se cierra, luego el script mide y reabre otro formulario

    def Vano2Click(self, sender, args):
        self.action = "vano2"
        self.Close()

    def AplicarClick(self, sender, args):
        if self.cmbRebar.SelectedIndex < 0:
            forms.alert("Debes seleccionar un diámetro de barra.", title="Falta información")
            return
        if self.L1_v1 is None or self.L1_v2 is None:
            forms.alert("Debes definir ambos vanos antes de aplicar.", title="Falta información")
            return
        self.action = "aplicar"
        self.Close()

    def CerrarClick(self, sender, args):
        self.action = "cerrar"
        self.Close()


xaml_path = os.path.join(os.path.dirname(__file__), "ArmaduraSuperiorForm.xaml")

# Variables que se irán llenando en el ciclo
L1_v1 = None
L1_v2 = None
largo1_cm = None
largo2_cm = None
L1_cm = None
L2_cm = None
sel_index = -1
esp_text_saved = None

# Bucle de interacción: en cada iteración se crea una NUEVA ventana,
# se muestra, y luego se evalúa la acción solicitada.
while True:
    form = ArmaduraSuperiorForm(
        xaml_path,
        L1_v1=L1_v1,
        L1_v2=L1_v2,
        largo1_cm=largo1_cm,
        largo2_cm=largo2_cm,
        L1_cm=L1_cm,
        L2_cm=L2_cm,
        sel_index=sel_index,
        esp_text=esp_text_saved
    )

    form.ShowDialog()

    if form.action == "cerrar" or form.action is None:
        # Cerrar formulario y terminar el comando después de haber creado las barras deseadas
        break

    if form.action == "vano1":
        try:
            p1_v1, p2_v1 = seleccionar_puntos()
        except OperationCanceledException:
            # Si presionas ESC al dibujar, simplemente volvemos a abrir el formulario
            continue
        largo1 = ((p2_v1.X - p1_v1.X) ** 2 + (p2_v1.Y - p1_v1.Y) ** 2) ** 0.5
        L1_v1 = largo1 / 3.0
        largo1_cm = UnitUtils.ConvertFromInternalUnits(largo1, UnitTypeId.Centimeters)
        L1_cm = UnitUtils.ConvertFromInternalUnits(L1_v1, UnitTypeId.Centimeters)
        sel_index = form.cmbRebar.SelectedIndex
        esp_text_saved = form.txtEspaciamiento.Text
        continue

    if form.action == "vano2":
        try:
            p1_v2, p2_v2 = seleccionar_puntos()
        except OperationCanceledException:
            continue
        largo2 = ((p2_v2.X - p1_v2.X) ** 2 + (p2_v2.Y - p1_v2.Y) ** 2) ** 0.5
        L1_v2 = largo2 / 3.0
        largo2_cm = UnitUtils.ConvertFromInternalUnits(largo2, UnitTypeId.Centimeters)
        L2_cm = UnitUtils.ConvertFromInternalUnits(L1_v2, UnitTypeId.Centimeters)
        sel_index = form.cmbRebar.SelectedIndex
        esp_text_saved = form.txtEspaciamiento.Text
        continue

    if form.action == "aplicar":
        # Validar que existan longitudes de ambos vanos
        L1 = L1_v1
        L2 = L1_v2
        if L1 is None or L2 is None:
            forms.alert("Debes definir ambos vanos antes de aplicar.", title="Falta información")
            continue

        # Tipo de barra seleccionado
        rebar_selected_name = form.cmbRebar.SelectedItem
        rebar_selected = next((x for x, y in zip(rebar_types, rebar_types_names) if y == rebar_selected_name), None)
        if not rebar_selected:
            forms.alert("No se seleccionó ningún tipo de barra.", exitscript=True)

        rebar_fi = rebar_selected.get_Parameter(DB.BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble()  # Diámetro seleccionado

        # Espaciamiento desde el formulario (cm)
        esp_text = form.txtEspaciamiento.Text
        try:
            esp_cm = float(esp_text)
        except (ValueError, TypeError):
            forms.alert("Ingresa un espaciamiento válido en centímetros (por ejemplo 20).")
            continue

        if esp_cm <= 0:
            forms.alert("El espaciamiento debe ser mayor que cero.")
            continue

        esp = UnitUtils.ConvertToInternalUnits(esp_cm, UnitTypeId.Centimeters)

# ===================================================================================
# TRAZADO DE BARRA
# ===================================================================================
        try:
            p1, p2 = seleccionar_puntos()
        except OperationCanceledException:
            # Si cancela el recorrido, vuelve al formulario
            continue

        ang = math.atan2(p2.Y - p1.Y, p2.X - p1.X)
        p1 = DB.XYZ(p1.X + (rec_top + rebar_fi / 2) * math.cos(ang), p1.Y + (rec_top + rebar_fi / 2) * math.sin(ang), p1.Z)
        p2 = DB.XYZ(p2.X - (rec_top + rebar_fi / 2) * math.cos(ang), p2.Y - (rec_top + rebar_fi / 2) * math.sin(ang), p2.Z)
        L_recorrido = ((p2.X - p1.X) ** 2 + (p2.Y - p1.Y) ** 2) ** 0.5
        vdir = (p2 - p1).Normalize()

        Z = Z_sup - rec_top - rebar_fi / 2
        p_start = DB.XYZ(p1.X - L1 * math.sin(ang), p1.Y + L1 * math.cos(ang), Z)
        p_end = DB.XYZ(p1.X + L2 * math.sin(ang), p1.Y - L2 * math.cos(ang), Z)
        curves = [DB.Line.CreateBound(p_start, p_end)]

 # ===================================================================================
# CREAR LA BARRA
# ===================================================================================
        with revit.Transaction("ARAINCO - Crear Rebar Set"):
            try:
                rebar = DB.Structure.Rebar.CreateFromCurves(
                    doc,
                    DB.Structure.RebarStyle.Standard,
                    rebar_selected,
                    None, None,
                    slab,
                    vdir,
                    curves,
                    DB.Structure.RebarHookOrientation.Right,
                    DB.Structure.RebarHookOrientation.Right,
                    True, True)

                N_barras = math.floor(L_recorrido / esp + 1)
                rebar.GetShapeDrivenAccessor().SetLayoutAsNumberWithSpacing(
                    N_barras,
                    esp,
                    True, True, True)
                rebar.SetPresentationMode(uidoc.ActiveView, DB.Structure.RebarPresentationMode.Middle)

                # Multi-Rebar Annotation
                nombre_tipo = "Recorrido Barras"
                tipos = DB.FilteredElementCollector(doc).OfClass(DB.MultiReferenceAnnotationType)
                tipo = next((t for t in tipos if DB.Element.Name.GetValue(t) == nombre_tipo), None)

                if tipo:
                    tipo_opts = DB.MultiReferenceAnnotationOptions(tipo)
                    rebar_ids = List[DB.ElementId]()
                    rebar_ids.Add(rebar.Id)
                    tipo_opts.SetElementsToDimension(rebar_ids)
                    tipo_opts.DimensionLineDirection = vdir
                    tipo_opts.DimensionPlaneNormal = uidoc.ActiveView.ViewDirection
                    offset = UnitUtils.ConvertToInternalUnits(80, UnitTypeId.Centimeters)
                    mid_point = DB.XYZ((p1.X + p2.X) / 2 - offset * math.sin(ang),
                                       (p1.Y + p2.Y) / 2 + offset * math.cos(ang),
                                       curves[0].GetEndPoint(0).Z)
                    tipo_opts.DimensionLineOrigin = mid_point
                    tipo_opts.TagHeadPosition = mid_point
                    DB.MultiReferenceAnnotation.Create(doc, uidoc.ActiveView.Id, tipo_opts)
                else:
                    forms.alert("La armadura se creó, pero no se encontró el tipo de cota llamado '{}'.\n\nRevisa que esté cargado en el proyecto o que el nombre esté escrito exactamente igual.".format(nombre_tipo), title="Falta Tipo de Recorrido")

                # Tag
                nombre_tag = "R_S_00"
                tipos = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).OfCategory(DB.BuiltInCategory.OST_RebarTags)
                tipo = next((t for t in tipos if DB.Element.Name.GetValue(t) == nombre_tag), None)

                subelementos = list(rebar.GetSubelements())
                idx_central = int(len(subelementos) / 2)
                rebar_central = subelementos[idx_central]
                centro_primera = DB.XYZ((p_start.X + p_end.X) / 2, (p_start.Y + p_end.Y) / 2, Z)
                centro_barra = DB.XYZ(centro_primera.X + (vdir.X * idx_central * esp),
                                      centro_primera.Y + (vdir.Y * idx_central * esp),
                                      Z)

                if tipo:
                    tag = DB.IndependentTag.Create(
                        doc,
                        tipo.Id,
                        uidoc.ActiveView.Id,
                        rebar_central.GetReference(),
                        False,
                        DB.TagOrientation.AnyModelDirection,
                        centro_barra)
                    tag.RotationAngle = ang + (math.pi / 2)
                else:
                    forms.alert("La armadura se creó, pero no se encontró el tipo de etiqueta llamado '{}'.\n\nRevisa que esté cargado en el proyecto o que el nombre esté escrito exactamente igual.".format(nombre_tag), title="Falta Etiqueta")

            except Exception as e:
                print("Error al crear barra, flecha de recorrido o tag: {}".format(e))

        # Después de aplicar un recorrido, volvemos al inicio del bucle para permitir más recorridos
        sel_index = form.cmbRebar.SelectedIndex
        esp_text_saved = form.txtEspaciamiento.Text
        continue
