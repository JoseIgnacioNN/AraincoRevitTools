# -*- coding: utf-8 -*-
import math
import os
import ctypes

import clr
clr.AddReference("PresentationFramework")

from Autodesk.Revit import DB
from Autodesk.Revit.DB import UnitUtils, UnitTypeId
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, forms

from System import EventHandler
from System.Collections.Generic import List


def get_z_on_plane(x, y, origen, normal):
    return origen.Z - (normal.X * (x - origen.X) + normal.Y * (y - origen.Y)) / normal.Z


class ScriptAController(object):
    """Controlador para Armadura Superior (script_A)."""

    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.view = uidoc.ActiveView

        if self.view.ViewType not in [DB.ViewType.EngineeringPlan, DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan]:
            forms.alert("Esta herramienta solo puede ser utilizada en vistas de planta.", title="Vista no válida", exitscript=True)

        self.Z_elev = self._get_level_elevation()
        self._ensure_work_plane()

        # Familia de línea para seleccionar vano
        nombre_familia = "EST_D_DEATIL ITEM_DIRECCION VANO MENOR"
        tipos = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements()
        self.tipo_linea = next((tip for tip in tipos if hasattr(tip, "Family") and tip.Family.Name == nombre_familia), None)
        if not self.tipo_linea:
            forms.alert(
                "No se encontró la familia '{}' en el proyecto.\nPor favor, cárgala antes de usar esta herramienta.".format(nombre_familia),
                exitscript=True,
            )

        # Tipos de barra
        self.rebar_types = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.Structure.RebarBarType)
            .WhereElementIsElementType()
            .ToElements()
        )
        self.rebar_types_names = [rt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for rt in self.rebar_types]

        # Estado interno
        self.largo1 = None
        self.largo2 = None
        self.L1 = None
        self.L2 = None
        self.sel_index = -1
        self.esp_text = None

    def _get_level_elevation(self):
        try:
            view_range = self.view.GetViewRange()
            top_level_id = view_range.GetLevelId(DB.PlanViewPlane.TopClipPlane)
            if top_level_id != DB.ElementId.InvalidElementId:
                return self.doc.GetElement(top_level_id).Elevation
            return self.view.GenLevel.Elevation
        except Exception:
            return self.view.GenLevel.Elevation

    def _ensure_work_plane(self):
        sp = self.view.SketchPlane
        needs = (
            (not sp)
            or abs(abs(sp.GetPlane().Normal.Z) - 1.0) > 0.001
            or abs(sp.GetPlane().Origin.Z - self.Z_elev) > 0.001
        )
        if needs:
            with revit.Transaction("ARAINCO - Configurar Work Plane"):
                plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, self.Z_elev))
                self.view.SketchPlane = DB.SketchPlane.Create(self.doc, plane)

    def seleccionar_puntos(self):
        teclas_enviadas = [False]
        id_capturado = []

        def capturador(sender, args):
            if args.GetAddedElementIds().Count > 0 and not teclas_enviadas[0]:
                id_capturado.append(args.GetAddedElementIds()[0])
                teclas_enviadas[0] = True

                VK_ESCAPE = 0x1B
                KEYEVENTF_KEYUP = 0x0002
                ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, 0, 0)
                ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)
                ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, 0, 0)
                ctypes.windll.user32.keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)

        t_ghost = DB.TransactionGroup(self.doc, "Borrador de Rastro")
        t_ghost.Start()

        self.doc.Application.DocumentChanged += EventHandler[DB.Events.DocumentChangedEventArgs](capturador)
        try:
            self.uidoc.PromptForFamilyInstancePlacement(self.tipo_linea)
        except OperationCanceledException as e:
            error_cancelacion = e
        finally:
            self.doc.Application.DocumentChanged -= EventHandler[DB.Events.DocumentChangedEventArgs](capturador)

        active_uiview = next((u for u in self.uidoc.GetOpenUIViews() if u.ViewId == self.uidoc.ActiveView.Id), None)
        zoom_corners = active_uiview.GetZoomCorners() if active_uiview else None

        if not id_capturado:
            t_ghost.RollBack()
            if zoom_corners:
                active_uiview.ZoomAndCenterRectangle(zoom_corners[0], zoom_corners[1])
            raise error_cancelacion or Exception("Operación cancelada")

        linea_id = id_capturado[0]
        p1 = self.doc.GetElement(linea_id).Location.Curve.GetEndPoint(0)
        p2 = self.doc.GetElement(linea_id).Location.Curve.GetEndPoint(1)
        t_ghost.RollBack()
        if zoom_corners:
            active_uiview.ZoomAndCenterRectangle(zoom_corners[0], zoom_corners[1])

        return p1, p2

    # ---------- Enlace con formulario ----------

    def bind_to_form(self, form):
        """Carga datos y combos en el formulario."""
        form.cmbRebar.ItemsSource = self.rebar_types_names
        if 0 <= self.sel_index < len(self.rebar_types_names):
            form.cmbRebar.SelectedIndex = self.sel_index

        if self.largo1 is not None:
            form.txtLargo1.Text = "{:.0f}".format(self.largo1)
        if self.L1 is not None:
            form.txtL1.Text = "{:.0f}".format(self.L1)
        if self.largo2 is not None:
            form.txtLargo2.Text = "{:.0f}".format(self.largo2)
        if self.L2 is not None:
            form.txtL2.Text = "{:.0f}".format(self.L2)
        if self.esp_text:
            form.txtEspaciamiento.Text = self.esp_text

    def save_from_form(self, form):
        self.sel_index = form.cmbRebar.SelectedIndex
        self.esp_text = form.txtEspaciamiento.Text

        if form.txtL1.Text:
            try:
                self.L1 = float(form.txtL1.Text)
            except ValueError:
                pass
        if form.txtL2.Text:
            try:
                self.L2 = float(form.txtL2.Text)
            except ValueError:
                pass

    def on_vano1(self, form):
        self.save_from_form(form)
        p1, p2 = self.seleccionar_puntos()
        self.largo1 = UnitUtils.ConvertFromInternalUnits((p2 - p1).GetLength(), UnitTypeId.Millimeters)
        self.L1 = math.ceil(self.largo1 / 3 / 10) * 10
        self.bind_to_form(form)

    def on_vano2(self, form):
        self.save_from_form(form)
        p1, p2 = self.seleccionar_puntos()
        self.largo2 = UnitUtils.ConvertFromInternalUnits((p2 - p1).GetLength(), UnitTypeId.Millimeters)
        self.L2 = math.ceil(self.largo2 / 3 / 10) * 10
        self.bind_to_form(form)

    def on_aplicar(self, form):
        # Validaciones (igual que antes)
        if form.cmbRebar.SelectedIndex < 0:
            forms.alert("Debes seleccionar un diámetro de barra.", title="Falta información")
            return
        if not form.txtL1.Text or not form.txtL2.Text:
            forms.alert("Debes definir ambos vanos antes de aplicar.", title="Falta información")
            return
        try:
            L1_mm = float(form.txtL1.Text)
            L2_mm = float(form.txtL2.Text)
        except ValueError:
            forms.alert("Los valores de L1 y L2 deben ser números válidos (solo números, sin letras).", title="Valor inválido")
            return
        if not form.txtEspaciamiento.Text:
            forms.alert("Debes ingresar un espaciamiento.", title="Falta información")
            return
        try:
            esp_mm = float(form.txtEspaciamiento.Text)
            if esp_mm <= 0:
                forms.alert("El espaciamiento debe ser mayor que cero.", title="Valor inválido")
                return
        except ValueError:
            forms.alert("Ingresa un espaciamiento válido en milímetros.", title="Valor inválido")
            return

        self.save_from_form(form)

        L1 = UnitUtils.ConvertToInternalUnits(L1_mm, UnitTypeId.Millimeters)
        L2 = UnitUtils.ConvertToInternalUnits(L2_mm, UnitTypeId.Millimeters)
        rebar_selected = self.rebar_types[form.cmbRebar.SelectedIndex]
        rebar_fi = rebar_selected.get_Parameter(DB.BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble()
        esp = UnitUtils.ConvertToInternalUnits(esp_mm, UnitTypeId.Millimeters)

        # Seleccionar recorrido
        try:
            p1, p2 = self.seleccionar_puntos()
        except OperationCanceledException:
            return

        # Transaction Group
        t_group = DB.TransactionGroup(self.doc, "ARAINCO - Armadura Superior sobre muro")
        t_group.Start()

        try:
            self._aplicar_armadura(p1, p2, L1, L2, rebar_selected, rebar_fi, esp, t_group)
        except Exception as e:
            if t_group.HasStarted():
                t_group.RollBack()
            forms.alert("Error al crear la barra: {}".format(e), title="Error de creación de barra")

    def _aplicar_armadura(self, p1, p2, L1, L2, rebar_selected, rebar_fi, esp, t_group):
        # --- Autodetección losa/fundación (copiado del script original) ---
        filtro_combinado = DB.LogicalOrFilter(
            DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Floors),
            DB.ElementCategoryFilter(DB.BuiltInCategory.OST_StructuralFoundation),
        )
        mid_point = DB.XYZ((p1.X + p2.X) / 2, (p1.Y + p2.Y) / 2, p1.Z)
        pt_min = DB.XYZ(mid_point.X - 0.0164, mid_point.Y - 0.0164, self.Z_elev - 6.562)
        pt_max = DB.XYZ(mid_point.X + 0.0164, mid_point.Y + 0.0164, self.Z_elev + 6.562)
        filtro_pilar = DB.BoundingBoxIntersectsFilter(DB.Outline(pt_min, pt_max))

        candidatos = (
            DB.FilteredElementCollector(self.doc, self.uidoc.ActiveView.Id)
            .WherePasses(filtro_combinado)
            .WherePasses(filtro_pilar)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        if not candidatos:
            t_group.RollBack()
            forms.alert(
                "No se encontró ninguna losa visible en un rango de ±1.5 metros respecto al nivel de la vista actual. Armadura no creada.",
                title="Losa no encontrada",
            )
            return

        opt = DB.Options()
        opt.DetailLevel = DB.ViewDetailLevel.Coarse
        p_test = DB.XYZ(mid_point.X, mid_point.Y, self.Z_elev)
        dist_min = float("inf")
        slab = None
        plane_normal = None
        plane_origin = None

        for candidato in candidatos:
            if not candidato.get_Geometry(opt):
                continue

            for geom_obj in candidato.get_Geometry(opt):
                if isinstance(geom_obj, DB.GeometryInstance):
                    iterable_geom = geom_obj.GetInstanceGeometry()
                else:
                    iterable_geom = [geom_obj]

                for solid in iterable_geom:
                    if isinstance(solid, DB.Solid) and solid.Faces.Size > 0:
                        for face in solid.Faces:
                            if isinstance(face, DB.PlanarFace) and face.FaceNormal.Z > 0.5:
                                inter = face.Project(p_test)
                                if not inter:
                                    continue
                                z_teorico = get_z_on_plane(mid_point.X, mid_point.Y, face.Origin, face.FaceNormal)
                                dist = abs(z_teorico - self.Z_elev)
                                if dist < dist_min:
                                    dist_min = dist
                                    slab = candidato
                                    plane_normal = face.FaceNormal
                                    plane_origin = face.Origin
                                    break

        if not slab:
            t_group.RollBack()
            mensaje = (
                "No se pudo detectar un host válido, por alguna de las siguientes razones:\n"
                "• La armadura se encuentra sobre un shaft.\n"
                "• La losa no tiene caras superiores válidas.\n"
                "• La losa no tiene prioridad de unión por sobre la viga, columna o muro."
            )
            forms.alert(mensaje, title="Anfitrión no encontrado")
            return

        try:
            rec_top_param = slab.get_Parameter(DB.BuiltInParameter.CLEAR_COVER_TOP)
            rec_top = self.doc.GetElement(rec_top_param.AsElementId()).CoverDistance
        except Exception:
            t_group.RollBack()
            forms.alert("La losa detectada no tiene parámetros de recubrimiento estructural válidos.", title="Error de Parámetro")
            return

        # --- Geometría barra ---
        offset_lateral = rec_top + rebar_fi / 2
        if (p2 - p1).GetLength() <= (2 * offset_lateral):
            t_group.RollBack()
            forms.alert("El recorrido dibujado es demasiado corto para aplicar los recubrimientos laterales.", title="Recorrido muy corto")
            return

        v_trazo = (p2 - p1).Normalize()
        p1 = p1 + v_trazo * offset_lateral
        p2 = p2 - v_trazo * offset_lateral
        origen_desplazado = plane_origin - plane_normal * (rec_top + rebar_fi / 2)
        p1_z = get_z_on_plane(p1.X, p1.Y, origen_desplazado, plane_normal)
        p2_z = get_z_on_plane(p2.X, p2.Y, origen_desplazado, plane_normal)
        p1_3d = DB.XYZ(p1.X, p1.Y, p1_z)
        p2_3d = DB.XYZ(p2.X, p2.Y, p2_z)
        vdir = (p2_3d - p1_3d).Normalize()
        v_bar = plane_normal.CrossProduct(vdir).Normalize()
        start = p1_3d + v_bar * L1
        end = p1_3d - v_bar * L2
        curves = [DB.Line.CreateBound(start, end)]
        L_recorrido = (p2_3d - p1_3d).GetLength()

        # --- Crear Rebar ---
        t = DB.Transaction(self.doc, "ARAINCO - Crear Rebar Set")
        t.Start()
        try:
            rebar = DB.Structure.Rebar.CreateFromCurves(
                self.doc,
                DB.Structure.RebarStyle.Standard,
                rebar_selected,
                None,
                None,
                slab,
                vdir,
                curves,
                DB.Structure.RebarHookOrientation.Right,
                DB.Structure.RebarHookOrientation.Right,
                True,
                True,
            )

            rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(esp, L_recorrido, True, True, True)
            self.doc.Regenerate()

            subelementos = list(rebar.GetSubelements())
            if len(subelementos) > 2:
                esp_real = L_recorrido / (len(subelementos) - 1)
                rebar.SetPresentationMode(self.uidoc.ActiveView, DB.Structure.RebarPresentationMode.Middle)
            else:
                esp_real = 0
                rebar.SetPresentationMode(self.uidoc.ActiveView, DB.Structure.RebarPresentationMode.All)

            # Offsets para MRA/Tag
            ang = math.atan2(p2.Y - p1.Y, p2.X - p1.X)
            if L1 > L2:
                offset_mra = UnitUtils.ConvertToInternalUnits(-80, UnitTypeId.Centimeters)
                offset_tag = UnitUtils.ConvertToInternalUnits(175, UnitTypeId.Centimeters)
            else:
                offset_mra = UnitUtils.ConvertToInternalUnits(80, UnitTypeId.Centimeters)
                offset_tag = UnitUtils.ConvertToInternalUnits(-175, UnitTypeId.Centimeters)

            # Multi-Rebar Annotation
            nombre_tipo = "Recorrido Barras"
            tipos = DB.FilteredElementCollector(self.doc).OfClass(DB.MultiReferenceAnnotationType)
            tipo = next((tt for tt in tipos if DB.Element.Name.GetValue(tt) == nombre_tipo), None)
            if tipo:
                tipo_opts = DB.MultiReferenceAnnotationOptions(tipo)
                rebar_ids = List[DB.ElementId]()
                rebar_ids.Add(rebar.Id)
                tipo_opts.SetElementsToDimension(rebar_ids)

                v_bar_2d = DB.XYZ(v_bar.X, v_bar.Y, 0).Normalize()
                vdir_mra = DB.XYZ(-v_bar_2d.Y, v_bar_2d.X, 0)
                v_trazo_2d = DB.XYZ(p2.X - p1.X, p2.Y - p1.Y, 0).Normalize()
                if vdir_mra.DotProduct(v_trazo_2d) < 0:
                    vdir_mra = -vdir_mra

                tipo_opts.DimensionLineDirection = vdir_mra
                tipo_opts.DimensionPlaneNormal = self.uidoc.ActiveView.ViewDirection
                mid = DB.XYZ(
                    (p1.X + p2.X) / 2 - offset_mra * math.sin(ang),
                    (p1.Y + p2.Y) / 2 + offset_mra * math.cos(ang),
                    self.Z_elev,
                )
                tipo_opts.DimensionLineOrigin = mid
                tipo_opts.TagHeadPosition = mid
                DB.MultiReferenceAnnotation.Create(self.doc, self.uidoc.ActiveView.Id, tipo_opts)
            else:
                forms.alert(
                    "La armadura se creó, pero no se encontró el tipo de cota llamado '{}'.\n\nRevisa que esté cargado en el proyecto o que el nombre esté escrito exactamente igual.".format(nombre_tipo),
                    title="Falta Tipo de Recorrido",
                )

            # Tag
            nombre_tag = "Marca - Cantidad - Diametro - Espaciamiento"
            tipos = (
                DB.FilteredElementCollector(self.doc)
                .OfClass(DB.FamilySymbol)
                .OfCategory(DB.BuiltInCategory.OST_RebarTags)
            )
            tipo = next((tt for tt in tipos if DB.Element.Name.GetValue(tt) == nombre_tag), None)

            idx_central = (len(subelementos) - 1) // 2
            rebar_central = subelementos[idx_central]
            centro_primera = DB.XYZ((start.X + end.X) / 2, (start.Y + end.Y) / 2, (start.Z + end.Z) / 2)
            centro_barra = centro_primera + vdir * idx_central * esp_real
            tag_pos = DB.XYZ(
                centro_barra.X - offset_tag * math.sin(ang),
                centro_barra.Y + offset_tag * math.cos(ang),
                self.Z_elev,
            )

            if tipo:
                try:
                    tag = DB.IndependentTag.Create(
                        self.doc,
                        tipo.Id,
                        self.uidoc.ActiveView.Id,
                        rebar_central.GetReference(),
                        False,
                        DB.TagOrientation.AnyModelDirection,
                        tag_pos,
                    )
                    tag.RotationAngle = ang + (math.pi / 2)
                except Exception:
                    forms.alert("Se creó la armadura, pero no fue posible etiquetarla", title="Error de etiqueta")
            else:
                forms.alert(
                    "La armadura se creó, pero no se encontró el tipo de etiqueta llamado '{}'.\n\nRevisa que esté cargado en el proyecto o que el nombre esté escrito exactamente igual.".format(nombre_tag),
                    title="Falta Etiqueta",
                )

            # Parámetros
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

            t.Commit()
            t_group.Assimilate()
        except Exception:
            if t.HasStarted():
                t.RollBack()
            raise

