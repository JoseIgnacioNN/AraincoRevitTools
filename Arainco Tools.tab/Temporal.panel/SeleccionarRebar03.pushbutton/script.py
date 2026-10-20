# -*- coding: utf-8 -*-
"""
Herramienta pyRevit para selección y filtrado de Structural Rebars por Rebar Shape.
Arainco Ingeniería Estructural — Revit API 2025+ (.NET 8).
"""

import os
import sys
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("System.Windows.Forms")

from System.Windows import Visibility, WindowStartupLocation
from System.Windows.Media import BrushConverter
from System.Windows.Forms import Application
from System.Collections.Generic import List

from pyrevit import revit, DB, UI, forms

# Helpers para compatibilidad Revit API 2025+ (ElementId de 64 bits)
try:
    from ag_utils import get_element_id_value, get_selected_elements
except ImportError:
    def get_element_id_value(el_id):
        if hasattr(el_id, "Value"):
            return el_id.Value
        return el_id.IntegerValue

    def get_selected_elements(doc=None, uidoc=None):
        doc = doc or revit.doc
        uidoc = uidoc or revit.uidoc
        if not doc or not uidoc:
            return []
        return [doc.GetElement(el_id) for el_id in uidoc.Selection.GetElementIds() if doc.GetElement(el_id)]


def get_rebar_shape_info(rebar, doc):
    """
    Retorna una tupla (shape_id, shape_name) de un elemento Rebar o RebarInSystem.
    Inspecciona múltiples fuentes de Revit API para garantizar máxima robustez.
    """
    shape_id = None
    shape_name = ""

    # 1. Probar método GetShapeId() de Revit API
    if hasattr(rebar, "GetShapeId"):
        try:
            s_id = rebar.GetShapeId()
            if s_id and s_id != DB.ElementId.InvalidElementId:
                shape_id = s_id
        except Exception:
            pass

    # 2. Probar GetAllRebarShapeIds() (para rebars con múltiples formas o freeform)
    if not shape_id and hasattr(rebar, "GetAllRebarShapeIds"):
        try:
            shape_ids = rebar.GetAllRebarShapeIds()
            if shape_ids and shape_ids.Count > 0:
                shape_id = shape_ids[0]
        except Exception:
            pass

    # 3. Probar propiedad RebarShapeId (RebarInSystem)
    if not shape_id and hasattr(rebar, "RebarShapeId"):
        try:
            s_id = rebar.RebarShapeId
            if s_id and s_id != DB.ElementId.InvalidElementId:
                shape_id = s_id
        except Exception:
            pass

    # 4. Obtener nombre desde el elemento RebarShape
    if shape_id and shape_id != DB.ElementId.InvalidElementId:
        try:
            shape_elem = doc.GetElement(shape_id)
            if shape_elem:
                shape_name = shape_elem.Name
        except Exception:
            pass

    # 5. Si no se obtuvo nombre, buscar en parámetro BuiltInParameter.REBAR_SHAPE
    if not shape_name:
        try:
            p_shape = rebar.get_Parameter(DB.BuiltInParameter.REBAR_SHAPE)
            if p_shape:
                val_str = p_shape.AsValueString()
                if val_str:
                    shape_name = val_str
                else:
                    p_id = p_shape.AsElementId()
                    if p_id and p_id != DB.ElementId.InvalidElementId:
                        shape_id = p_id
                        shape_elem = doc.GetElement(p_id)
                        if shape_elem:
                            shape_name = shape_elem.Name
        except Exception:
            pass

    # 6. Fallback en parámetros por nombre común en español e inglés
    if not shape_name:
        for p_name in ["Shape", "Forma", "Forma de armadura", "Rebar Shape", "Shape Name"]:
            try:
                p = rebar.LookupParameter(p_name)
                if p:
                    val = p.AsString() or p.AsValueString()
                    if val:
                        shape_name = val
                        break
            except Exception:
                pass

    return shape_id, (shape_name.strip() if shape_name else "Sin Forma")


def shape_matches(shape_name, target_text, match_exact=False):
    """Evalúa si el nombre de forma coincide con el texto objetivo."""
    if not shape_name or not target_text:
        return False
    s_name = shape_name.strip().lower()
    t_text = target_text.strip().lower()

    if match_exact:
        return s_name == t_text
    else:
        return t_text in s_name


class DarkAlertWindow(forms.WPFWindow):
    """Diálogo modal corporativo Arainco para advertencias y avisos."""

    def __init__(self, message, title="Arainco — Aviso", header=None, status="info"):
        xaml_path = os.path.join(os.path.dirname(__file__), "alert_ui.xaml")
        forms.WPFWindow.__init__(self, xaml_path)
        self.Title = title
        self.lbl_header.Text = header or title.replace("Arainco — ", "")
        self.lbl_message.Text = message

        converter = BrushConverter()
        if status == "success":
            accent = converter.ConvertFromString("#4ECDC4")
            self.lbl_icon.Text = "✓"
        elif status == "warning":
            accent = converter.ConvertFromString("#F3C65F")
            self.lbl_icon.Text = "!"
        elif status == "danger":
            accent = converter.ConvertFromString("#E06C75")
            self.lbl_icon.Text = "✕"
        else:
            accent = converter.ConvertFromString("#5DB9EA")
            self.lbl_icon.Text = "i"

        self.bd_badge.BorderBrush = accent
        self.lbl_icon.Foreground = accent

    def on_ok_click(self, sender, args):
        self.Close()


def show_alert(message, title="Arainco — Aviso", header=None, status="info"):
    """Muestra cuadro de diálogo modal con paleta Arainco."""
    dlg = DarkAlertWindow(message=message, title=title, header=header, status=status)
    dlg.ShowDialog()


class DarkSummaryWindow(forms.WPFWindow):
    """Diálogo modal corporativo Arainco para resumen de ejecución."""

    def __init__(self, total_rebars, seleccionados, formas_count, target_shape, detalles, is_success=True):
        xaml_path = os.path.join(os.path.dirname(__file__), "summary_ui.xaml")
        forms.WPFWindow.__init__(self, xaml_path)
        self.Title = "Arainco — Selección de Rebars: Resumen"

        converter = BrushConverter()
        accent = converter.ConvertFromString("#4ECDC4" if is_success else "#F3C65F")
        self.bd_badge.BorderBrush = accent
        self.lbl_status_icon.Foreground = accent
        self.lbl_status_icon.Text = "✓ Selección Activa en Revit" if is_success else "! Sin coincidencias"

        self.lbl_kpi_total.Text = str(total_rebars)
        self.lbl_kpi_seleccionados.Text = str(seleccionados)
        self.lbl_kpi_formas.Text = str(formas_count)
        self.lbl_kpi_shape_label.Text = "Con Rebar Shape '{}'".format(target_shape)
        self.lbl_summary_details.Text = detalles

    def on_cerrar_click(self, sender, args):
        self.Close()


class SeleccionarRebar03Window(forms.WPFWindow):
    """Ventana principal WPF para filtrar y seleccionar rebars por forma."""

    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), "ui.xaml")
        forms.WPFWindow.__init__(self, xaml_path)

        self.doc = revit.doc
        self.uidoc = revit.uidoc
        self.view = self.doc.ActiveView
        self._is_updating_ui = False

        # Datos de la vista activa
        self.lbl_vista_nombre.Text = self.view.Name
        self.lbl_vista_escala.Text = "1:{} — {}".format(self.view.Scale, str(self.view.ViewType))

        # 1. Obtener rebars de la vista activa
        col_view = (DB.FilteredElementCollector(self.doc, self.view.Id)
                    .OfCategory(DB.BuiltInCategory.OST_Rebar)
                    .WhereElementIsNotElementType())
        self.all_view_rebars = list(col_view)

        # 2. Evaluar selección previa
        selected_elements = get_selected_elements(self.doc, self.uidoc)
        self.pre_selected_rebars = []
        for el in selected_elements:
            try:
                if el.Category and get_element_id_value(el.Category.Id) == int(DB.BuiltInCategory.OST_Rebar):
                    self.pre_selected_rebars.append(el)
            except Exception:
                pass

        # Configurar radio buttons de ámbito
        if self.pre_selected_rebars:
            self.rb_ambito_seleccion.IsEnabled = True
            self.rb_ambito_seleccion.Content = "Solo dentro de selección previa ({} rebars)".format(len(self.pre_selected_rebars))
        else:
            self.rb_ambito_seleccion.IsEnabled = False
            self.rb_ambito_seleccion.Content = "Solo dentro de selección previa (0 rebars)"

        # 3. Analizar formas presentes en los rebars de la vista
        self._analizar_rebars_y_formas()

    def _get_target_rebars_pool(self):
        """Retorna la lista de rebars a evaluar según el ámbito seleccionado."""
        if self.rb_ambito_seleccion.IsChecked and self.pre_selected_rebars:
            return self.pre_selected_rebars
        return self.all_view_rebars

    def _analizar_rebars_y_formas(self):
        """Escanea los rebars del ámbito actual para catalogar formas y frecuencias."""
        self._is_updating_ui = True
        rebars_pool = self._get_target_rebars_pool()

        self.shape_counts = {}  # nombre_forma -> int
        self.rebar_shapes_map = {}  # rebar_id_val -> nombre_forma

        for r in rebars_pool:
            _, s_name = get_rebar_shape_info(r, self.doc)
            r_val = get_element_id_value(r.Id)
            self.rebar_shapes_map[r_val] = s_name
            self.shape_counts[s_name] = self.shape_counts.get(s_name, 0) + 1

        # Poblar ComboBox de formas detectadas
        self.cb_shapes_disponibles.Items.Clear()
        self.cb_shapes_disponibles.Items.Add("-- Seleccionar de la lista --")

        # Ordenar formas por frecuencia (de mayor a menor)
        sorted_shapes = sorted(self.shape_counts.items(), key=lambda item: item[1], reverse=True)
        selected_combo_idx = 0

        target_text = self.txt_shape_filter.Text.strip().lower() if hasattr(self, "txt_shape_filter") else "03"
        for idx, (s_name, count) in enumerate(sorted_shapes):
            item_text = "{} ({} rebars)".format(s_name, count)
            self.cb_shapes_disponibles.Items.Add(item_text)
            if target_text and target_text in s_name.lower() and selected_combo_idx == 0:
                selected_combo_idx = idx + 1

        if selected_combo_idx > 0:
            self.cb_shapes_disponibles.SelectedIndex = selected_combo_idx
        else:
            self.cb_shapes_disponibles.SelectedIndex = 0

        self._is_updating_ui = False
        self._recalcular_kpis()

    def _recalcular_kpis(self):
        """Actualiza las tarjetas de KPIs en tiempo real según el filtro activo."""
        rebars_pool = self._get_target_rebars_pool()
        total_pool = len(rebars_pool)
        target_text = self.txt_shape_filter.Text.strip() if hasattr(self, "txt_shape_filter") else "03"
        match_exact = self.rb_match_exact.IsChecked if hasattr(self, "rb_match_exact") else False

        matching_count = 0
        for r in rebars_pool:
            r_val = get_element_id_value(r.Id)
            s_name = self.rebar_shapes_map.get(r_val, "")
            if shape_matches(s_name, target_text, match_exact):
                matching_count += 1

        self.lbl_kpi_vista_total.Text = str(total_pool)
        self.lbl_kpi_vista_forma03.Text = str(matching_count)
        self.lbl_kpi_vista_formas_count.Text = str(len(self.shape_counts))

        # Habilitar o atenuar botón de ejecutar
        self.btn_ejecutar.IsEnabled = (total_pool > 0)

    def on_ambito_changed(self, sender, args):
        """Maneja el cambio de radio button entre vista activa y selección previa."""
        if not self._is_updating_ui:
            self._analizar_rebars_y_formas()

    def on_filter_changed(self, sender, args):
        """Maneja cambios en el texto del filtro o tipo de coincidencia."""
        if not self._is_updating_ui:
            self._recalcular_kpis()

    def on_shape_selected(self, sender, args):
        """Maneja la selección de un ítem en el ComboBox de formas detectadas."""
        if self._is_updating_ui:
            return

        selected_item = self.cb_shapes_disponibles.SelectedItem
        if selected_item and selected_item != "-- Seleccionar de la lista --":
            # Extraer el nombre de la forma descartando la parte de ' (XX rebars)'
            shape_name_part = str(selected_item).split(" (")[0].strip()
            self._is_updating_ui = True
            self.txt_shape_filter.Text = shape_name_part
            self._is_updating_ui = False
            self._recalcular_kpis()

    def on_manual_click(self, sender, args):
        """Abre el manual interactivo de usuario HTML."""
        manual_path = os.path.join(os.path.dirname(__file__), "manual_usuario.html")
        if os.path.exists(manual_path):
            os.startfile(manual_path)
        else:
            show_alert("No se encontró el manual de usuario en:\n{}".format(manual_path),
                       title="Arainco — Manual no encontrado", status="warning")

    def on_cancelar_click(self, sender, args):
        """Cierra la ventana sin realizar acciones."""
        self.Close()

    def _actualizar_progreso(self, actual, total, texto):
        """Actualiza la barra de progreso WPF en tiempo real."""
        porcentaje = int((float(actual) / float(total)) * 100.0) if total > 0 else 100
        self.pb_ejecucion.Value = porcentaje
        self.lbl_progress_percent.Text = "{}%".format(porcentaje)
        self.lbl_progress_status.Text = texto
        Application.DoEvents()

    def on_ejecutar_click(self, sender, args):
        """Ejecuta el filtrado y selección de los rebars que coinciden con el Shape."""
        target_text = self.txt_shape_filter.Text.strip()
        if not target_text:
            show_alert("Por favor ingrese un nombre o código de Rebar Shape a filtrar.",
                       title="Arainco — Parámetro Vacío", status="warning")
            return

        rebars_pool = self._get_target_rebars_pool()
        if not rebars_pool:
            show_alert("No se encontraron armaduras (Structural Rebars) en el ámbito seleccionado.",
                       title="Arainco — Sin Elementos", status="warning")
            return

        match_exact = self.rb_match_exact.IsChecked
        aislar_en_vista = self.chk_aislar_vista.IsChecked
        zoom_seleccion = self.chk_zoom_seleccion.IsChecked

        # Mostrar barra de progreso
        self.bd_progress.Visibility = Visibility.Visible
        self._actualizar_progreso(0, len(rebars_pool), "Iniciando análisis de Rebars...")

        matching_rebars = []
        element_ids_to_select = List[DB.ElementId]()
        formas_coincidentes_map = {}

        total_count = len(rebars_pool)
        for idx, rebar in enumerate(rebars_pool):
            r_val = get_element_id_value(rebar.Id)
            s_name = self.rebar_shapes_map.get(r_val, "")
            if shape_matches(s_name, target_text, match_exact):
                matching_rebars.append(rebar)
                element_ids_to_select.Add(rebar.Id)
                formas_coincidentes_map[s_name] = formas_coincidentes_map.get(s_name, 0) + 1

            if idx % 15 == 0 or idx == total_count - 1:
                self._actualizar_progreso(
                    idx + 1, total_count,
                    "Analizando rebar {} de {} ({} coincidentes)...".format(idx + 1, total_count, len(matching_rebars))
                )

        if not matching_rebars:
            self.bd_progress.Visibility = Visibility.Collapsed
            show_alert(
                "No se encontró ningún Structural Rebar cuya forma coincida con '{}' (Modo: {}).\n\n"
                "Formas presentes en la vista:\n{}".format(
                    target_text,
                    "Exacta" if match_exact else "Contiene",
                    ", ".join(["'{}' ({})".format(k, v) for k, v in sorted(self.shape_counts.items())])
                ),
                title="Arainco — Sin Coincidencias",
                status="warning"
            )
            return

        # 1. Aplicar selección en Revit UI
        try:
            self.uidoc.Selection.SetElementIds(element_ids_to_select)
        except Exception as ex_sel:
            self.bd_progress.Visibility = Visibility.Collapsed
            show_alert("Error al establecer la selección en Revit:\n{}".format(str(ex_sel)),
                       title="Arainco — Error de Selección", status="danger")
            return

        # 2. Si se solicitó aislar temporalmente en la vista
        if aislar_en_vista:
            try:
                with revit.Transaction("Arainco: Aislar Rebars Forma {}".format(target_text)):
                    self.view.IsolateElementsTemporary(element_ids_to_select)
            except Exception as ex_iso:
                # El aislamiento de vista no debe interrumpir el flujo si la vista no lo permite
                pass

        # 3. Si se solicitó zoom a la selección
        if zoom_seleccion:
            try:
                self.uidoc.ShowElements(element_ids_to_select)
            except Exception:
                pass

        # Cerrar ventana principal y presentar resumen
        self.Close()

        # Construir reporte detallado
        detalles_lineas = [
            "Filtro aplicado: '{}' (Modo: {})".format(target_text, "Exacto" if match_exact else "Contiene"),
            "Ámbito analizado: {}".format("Selección previa" if self.rb_ambito_seleccion.IsChecked else "Toda la vista activa"),
            "Vista: {} (1:{})".format(self.view.Name, self.view.Scale),
            "",
            "Desglose de formas coincidentes seleccionadas:"
        ]
        for s_n, c_count in sorted(formas_coincidentes_map.items()):
            detalles_lineas.append("  • Forma '{}': {} armaduras".format(s_n, c_count))

        detalles_lineas.append("")
        detalles_lineas.append("Acciones complementarias:")
        detalles_lineas.append("  • Aislamiento temporal en vista: {}".format("ACTIVADO" if aislar_en_vista else "No aplicado"))
        detalles_lineas.append("  • Zoom a la selección: {}".format("APLICADO" if zoom_seleccion else "No aplicado"))

        detalles_lineas.append("")
        detalles_lineas.append("IDs de elementos seleccionados (primeros 50):")
        ids_preview = [str(get_element_id_value(r.Id)) for r in matching_rebars[:50]]
        detalles_lineas.append(", ".join(ids_preview))
        if len(matching_rebars) > 50:
            detalles_lineas.append("... y {} armaduras más.".format(len(matching_rebars) - 50))

        detalles_texto = "\n".join(detalles_lineas)

        summary_dlg = DarkSummaryWindow(
            total_rebars=total_count,
            seleccionados=len(matching_rebars),
            formas_count=len(self.shape_counts),
            target_shape=target_text,
            detalles=detalles_texto,
            is_success=True
        )
        summary_dlg.ShowDialog()


# Punto de entrada de la herramienta
if __name__ == "__main__":
    if not revit.doc or not revit.uidoc:
        show_alert("No hay un proyecto activo abierto en Revit.",
                   title="Arainco — Documento no disponible", status="warning")
    elif revit.active_view is None:
        show_alert("No hay una vista activa en Revit.",
                   title="Arainco — Vista no disponible", status="warning")
    elif revit.active_view.ViewType == DB.ViewType.DrawingSheet:
        show_alert(
            "La vista activa es una Lámina (Sheet).\n\n"
            "Por favor abra o active una vista de modelo (Planta, Elevación, Sección o 3D) "
            "donde se encuentren visibles las armaduras estructurales (Structural Rebars).",
            title="Arainco — Vista no válida",
            status="warning"
        )
    else:
        win = SeleccionarRebar03Window()
        win.ShowDialog()
