# -*- coding: utf-8 -*-
"""Controlador WPF del wizard de armado de columnas v2.

Patrón:  WindowController (imperativo) + WizardViewModel (estado puro).

La lógica de negocio vive en el ViewModel; este archivo solo:
 - Carga el XAML
 - Convierte eventos WPF → llamadas al ViewModel
 - Lee el ViewModel y actualiza controles WPF directamente
 - Dibuja en los Canvas
"""

import os

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import AppDomain
from System.IO import File
from System.Windows import (
    Visibility, Thickness,
    HorizontalAlignment as HAlign,
    VerticalAlignment as VAlign,
    FontWeights as FW,
    CornerRadius,
    GridLength, GridUnitType,
    TextAlignment,
)
from System.Windows.Controls import (
    StackPanel, Border, Grid, TextBlock, CheckBox, ComboBox, Button,
    ComboBoxItem, ScrollViewer, RowDefinition, ColumnDefinition,
    Orientation, Canvas as WpfCanvas,
)
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Shapes import Rectangle as WpfRect, Ellipse
from System.Windows.Markup import XamlReader

try:
    clr.AddReference("RevitAPIUI")
    from Autodesk.Revit.UI import TaskDialog
except Exception:
    TaskDialog = None

from column_reinforcement_v2.ui.wizard_viewmodel import WizardViewModel
from column_reinforcement_v2.models.splice_segment import SEGMENT_COLORS

_SINGLETON_KEY = "Arainco.column_reinforcement_v2.WizardWindowSingleton"

# Colores de grupo (badge)
GROUP_BADGE_COLORS = ["#1B6CA8", "#D4AC0D", "#E67E22", "#8E44AD", "#17A589", "#C0392B"]


def _hex_to_brush(hex_color):
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return SolidColorBrush(Color.FromRgb(r, g, b))


class WizardWindowController(object):
    """Orquesta la ventana WPF con el WizardViewModel."""

    AVAILABLE_DIAMETERS = [8, 10, 12, 16, 20, 22, 25, 28, 32, 36]

    def __init__(self, viewmodel=None):
        self.vm = viewmodel or WizardViewModel()
        self.request = None
        self._current_step = 1
        self._diameter_combos = {}   # segment_id → ComboBox
        self._bar_labels = {}        # (group_id, side) → TextBlock
        self._splice_checks = {}     # z_mm → CheckBox

        xaml_text = self._load_xaml()
        self.window = XamlReader.Parse(xaml_text)
        self._bind_named_controls()
        self._wire_events()

    # ------------------------------------------------------------------ #
    #  Carga XAML                                                         #
    # ------------------------------------------------------------------ #

    @staticmethod
    def _load_xaml():
        here = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(here, "main_window.xaml")
        return File.ReadAllText(path)

    def _find(self, name):
        return self.window.FindName(name)

    def _bind_named_controls(self):
        """Guarda referencias a controles nombrados para acceso rápido."""
        # Step 1
        self._canvas_building = self._find("BuildingPreviewCanvas")
        self._info_sel_count   = self._find("InfoSelCount")
        self._info_height      = self._find("InfoTotalHeight")
        self._info_from        = self._find("InfoFrom")
        self._info_to          = self._find("InfoTo")
        # Step 2
        self._splice_panel     = self._find("SpliceCheckPanel")
        self._info_seg_count   = self._find("InfoSegCount")
        self._info_lap         = self._find("InfoLapHeight")
        # Step 3
        self._groups_panel     = self._find("GroupsPanel")
        # Step 4
        self._segments_panel   = self._find("SegmentsPanel")
        self._canvas_preview   = self._find("SegmentPreviewCanvas")
        self._info_seg_gen     = self._find("InfoSegGenerated")
        # Footer
        self._footer_cols      = self._find("FooterColCount")
        self._footer_height    = self._find("FooterTotalHeight")
        self._footer_segs      = self._find("FooterSegments")
        # Step circles
        self._step_circles     = [
            self._find("StepCircle{0}".format(i)) for i in range(1, 5)
        ]
        self._step_labels      = [
            self._find("StepLabel{0}".format(i)) for i in range(1, 5)
        ]
        # Panel borders
        self._panel_borders    = [
            self._find("Panel{0}Border".format(i)) for i in range(1, 5)
        ]

    # ------------------------------------------------------------------ #
    #  Wiring de eventos                                                  #
    # ------------------------------------------------------------------ #

    def _wire_events(self):
        self._find("BtnClose").Click   += self._on_close
        self._find("BtnHelp").Click    += self._on_help
        self._find("BtnNext1").Click   += self._on_next1
        self._find("BtnBack2").Click   += self._on_back
        self._find("BtnNext2").Click   += self._on_next2
        self._find("BtnBack3").Click   += self._on_back
        self._find("BtnNext3").Click   += self._on_next3
        self._find("BtnBack4").Click   += self._on_back
        self._find("BtnGenerate").Click += self._on_generate
        self.window.Closed += self._on_closed

    # ------------------------------------------------------------------ #
    #  Carga inicial de datos                                             #
    # ------------------------------------------------------------------ #

    def populate(self, column_elements):
        """Carga los elementos Revit en el ViewModel y actualiza toda la UI."""
        self.vm.load_from_elements(column_elements)
        self._refresh_step1()
        self._build_splice_panel()
        self._refresh_step2()
        self._build_groups_panel()
        self._build_segments_panel()
        self._refresh_footer()
        self._set_active_step(1)

    # ------------------------------------------------------------------ #
    #  Actualización de la UI por paso                                    #
    # ------------------------------------------------------------------ #

    def _refresh_step1(self):
        vm = self.vm
        self._info_sel_count.Text = str(vm.total_columns)
        self._info_height.Text    = u"{0:.2f} m".format(vm.total_height_m)
        self._info_from.Text      = u"N+{0:.2f}".format(vm.z_bottom_mm / 1000.0)
        self._info_to.Text        = u"N+{0:.2f}".format(vm.z_top_mm / 1000.0)
        self._draw_building_preview()

    def _refresh_step2(self):
        self._info_seg_count.Text = str(self.vm.segment_count)
        self._info_lap.Text       = self.vm.average_lap_height_label

    def _refresh_footer(self):
        vm = self.vm
        self._footer_cols.Text   = str(vm.total_columns)
        self._footer_height.Text = u"{0:.2f} m".format(vm.total_height_m)
        self._footer_segs.Text   = str(vm.segment_count)
        self._info_seg_gen.Text  = str(vm.segment_count)

    # ------------------------------------------------------------------ #
    #  Dibujos en Canvas                                                  #
    # ------------------------------------------------------------------ #

    def _draw_building_preview(self):
        canvas = self._canvas_building
        canvas.Children.Clear()

        groups = self.vm.column_groups
        if not groups:
            return

        z_min = self.vm.z_bottom_mm
        z_max = self.vm.z_top_mm
        total_h = max(z_max - z_min, 1.0)

        # Dimensiones del canvas (aprox por el XAML Height="200")
        canvas_h = 190.0
        canvas_w = 120.0
        bar_w    = 30.0
        x_offset = (canvas_w - bar_w) / 2.0

        for idx, group in enumerate(groups):
            ratio_bot = (group.z_bottom_mm - z_min) / total_h
            ratio_top = (group.z_top_mm - z_min) / total_h
            y_bot = canvas_h - ratio_top * canvas_h
            h_px  = (ratio_top - ratio_bot) * canvas_h
            if h_px < 2:
                h_px = 2

            color_hex = GROUP_BADGE_COLORS[idx % len(GROUP_BADGE_COLORS)]
            rect = WpfRect()
            rect.Width  = bar_w
            rect.Height = h_px
            rect.Fill   = _hex_to_brush(color_hex)
            rect.Opacity = 0.85
            WpfCanvas.SetLeft(rect, x_offset)
            WpfCanvas.SetTop(rect, y_bot)
            canvas.Children.Add(rect)

            # Label de nivel
            lbl = TextBlock()
            lbl.Text     = u"N+{0:.1f}".format(group.z_top_mm / 1000.0)
            lbl.FontSize = 9
            lbl.Foreground = _hex_to_brush("#95B8CC")
            WpfCanvas.SetLeft(lbl, x_offset + bar_w + 4)
            WpfCanvas.SetTop(lbl, y_bot - 6)
            canvas.Children.Add(lbl)

        # Nivel base
        lbl_base = TextBlock()
        lbl_base.Text     = u"N+{0:.1f}".format(z_min / 1000.0)
        lbl_base.FontSize = 9
        lbl_base.Foreground = _hex_to_brush("#95B8CC")
        WpfCanvas.SetLeft(lbl_base, x_offset + bar_w + 4)
        WpfCanvas.SetTop(lbl_base, canvas_h - 14)
        canvas.Children.Add(lbl_base)

    def _draw_segment_preview(self):
        canvas = self._canvas_preview
        canvas.Children.Clear()

        segs = self.vm.splice_segments
        if not segs:
            return

        z_min = min(s.z_start_mm for s in segs)
        z_max = max(s.z_end_mm   for s in segs)
        total_h = max(z_max - z_min, 1.0)
        canvas_h = 85.0
        bar_w    = 18.0

        # Una columna de preview por segmento (lado a lado)
        seg_w  = 16.0
        gap    = 3.0
        x_base = 10.0

        for idx, seg in enumerate(segs):
            ratio_bot = (seg.z_start_mm - z_min) / total_h
            ratio_top = (seg.z_end_mm   - z_min) / total_h
            y_bot = canvas_h - ratio_top * canvas_h
            h_px  = max((ratio_top - ratio_bot) * canvas_h, 2.0)
            x     = x_base

            rect = WpfRect()
            rect.Width  = bar_w
            rect.Height = h_px
            rect.Fill   = _hex_to_brush(seg.color)
            rect.Opacity = 0.85
            WpfCanvas.SetLeft(rect, x)
            WpfCanvas.SetTop(rect, y_bot)
            canvas.Children.Add(rect)

            # Badge número
            lbl = TextBlock()
            lbl.Text     = str(seg.segment_id)
            lbl.FontSize = 8
            lbl.FontWeight = _fw_semibold()
            lbl.Foreground = _hex_to_brush("#E8F4F8")
            WpfCanvas.SetLeft(lbl, x + 5)
            WpfCanvas.SetTop(lbl, y_bot + h_px / 2.0 - 6)
            canvas.Children.Add(lbl)

        # Leyenda de niveles a la derecha
        x_legend = x_base + bar_w + 8
        for seg in segs:
            ratio_top = (seg.z_end_mm - z_min) / total_h
            y_top = canvas_h - ratio_top * canvas_h
            lbl = TextBlock()
            lbl.Text     = u"N+{0:.2f}".format(seg.z_end_mm / 1000.0)
            lbl.FontSize = 8
            lbl.Foreground = _hex_to_brush("#95B8CC")
            WpfCanvas.SetLeft(lbl, x_legend)
            WpfCanvas.SetTop(lbl, max(y_top - 6, 0))
            canvas.Children.Add(lbl)

    # ------------------------------------------------------------------ #
    #  Construcción dinámica de paneles                                   #
    # ------------------------------------------------------------------ #

    def _build_splice_panel(self):
        panel = self._splice_panel
        panel.Children.Clear()
        self._splice_checks.clear()

        for z_mm, label, is_cut, is_active in self.vm.all_cut_points_sorted:
            row = StackPanel()
            row.Orientation = Orientation.Horizontal

            if is_cut:
                cb = CheckBox()
                cb.Content   = label
                cb.IsChecked = is_active
                cb.Margin    = Thickness(0, 3, 0, 3)
                cb.Tag       = z_mm
                cb.Checked   += self._on_splice_check_changed
                cb.Unchecked += self._on_splice_check_changed
                self._splice_checks[z_mm] = cb
                row.Children.Add(cb)
            else:
                lbl = TextBlock()
                lbl.Text     = label
                lbl.Foreground = _hex_to_brush("#4A7A94")
                lbl.Margin   = Thickness(20, 4, 0, 4)
                lbl.FontSize = 10
                row.Children.Add(lbl)

            panel.Children.Add(row)

    def _build_groups_panel(self):
        panel = self._groups_panel
        panel.Children.Clear()
        self._bar_labels.clear()

        for group in self.vm.column_groups:
            dist = self.vm.distribution_for(group.group_id)
            if dist is None:
                continue

            color_hex = GROUP_BADGE_COLORS[(group.group_id - 1) % len(GROUP_BADGE_COLORS)]

            row = Grid()
            row.Margin = Thickness(0, 3, 0, 3)

            # Columnas del grid interno
            for w in [32, -1, 55, 60, 60]:  # -1 = star
                cd = ColumnDefinition()
                if w == -1:
                    cd.Width = _star_width()
                else:
                    cd.Width = _pixel_width(w)
                row.ColumnDefinitions.Add(cd)

            # Badge grupo
            badge = Border()
            badge.Background    = _hex_to_brush(color_hex)
            badge.CornerRadius  = _corner_radius(12)
            badge.Width, badge.Height = 24, 24
            badge.HorizontalAlignment = HAlign.Center
            badge.VerticalAlignment   = VAlign.Center
            num_lbl = TextBlock()
            num_lbl.Text       = str(group.group_id)
            num_lbl.FontWeight = FW.Bold
            num_lbl.FontSize   = 11
            num_lbl.Foreground = _hex_to_brush("#E8F4F8")
            num_lbl.HorizontalAlignment = HAlign.Center
            num_lbl.VerticalAlignment   = VAlign.Center
            badge.Child = num_lbl
            _set_col(badge, 0)
            row.Children.Add(badge)

            # Sección + rango
            sec_panel = StackPanel()
            sec_lbl = TextBlock()
            sec_lbl.Text     = group.section.label()
            sec_lbl.FontSize = 11
            sec_lbl.FontWeight = _fw_semibold()
            sec_lbl.Foreground = _hex_to_brush("#E8F4F8")
            rng_lbl = TextBlock()
            rng_lbl.Text     = group.column_range_label()
            rng_lbl.FontSize = 9
            rng_lbl.Foreground = _hex_to_brush("#95B8CC")
            sec_panel.Children.Add(sec_lbl)
            sec_panel.Children.Add(rng_lbl)
            _set_col(sec_panel, 1)
            row.Children.Add(sec_panel)

            # Contador de columnas
            col_count_lbl = TextBlock()
            col_count_lbl.Text = str(group.column_count)
            col_count_lbl.FontSize = 11
            col_count_lbl.HorizontalAlignment = HAlign.Center
            col_count_lbl.VerticalAlignment   = VAlign.Center
            col_count_lbl.Foreground = _hex_to_brush("#95B8CC")
            _set_col(col_count_lbl, 2)
            row.Children.Add(col_count_lbl)

            # Spinner A
            spin_a = self._make_spinner(group.group_id, "A", dist.side_a_count)
            _set_col(spin_a, 3)
            row.Children.Add(spin_a)

            # Spinner B
            spin_b = self._make_spinner(group.group_id, "B", dist.side_b_count)
            _set_col(spin_b, 4)
            row.Children.Add(spin_b)

            panel.Children.Add(row)

            # Separador
            sep = WpfRect()
            sep.Height  = 1
            sep.Fill    = _hex_to_brush("#1A3040")
            sep.Margin  = Thickness(0, 2, 0, 2)
            panel.Children.Add(sep)

    def _make_spinner(self, group_id, side, initial_value):
        """Crea un control +/- para el número de barras."""
        container = StackPanel()
        container.Orientation = Orientation.Horizontal
        container.HorizontalAlignment = HAlign.Center
        container.VerticalAlignment   = VAlign.Center

        btn_minus = Button()
        btn_minus.Content = u"−"
        btn_minus.Style   = self.window.Resources["BtnSpin"]
        btn_minus.Tag     = "{0}:{1}:dec".format(group_id, side)
        btn_minus.Click  += self._on_spin_click

        lbl = TextBlock()
        lbl.Text              = str(initial_value)
        lbl.Width             = 22
        lbl.FontSize          = 12
        lbl.FontWeight        = FW.SemiBold
        lbl.Foreground        = _hex_to_brush("#E8F4F8")
        lbl.HorizontalAlignment = HAlign.Center
        lbl.VerticalAlignment   = VAlign.Center
        lbl.TextAlignment       = TextAlignment.Center
        self._bar_labels[(group_id, side)] = lbl

        btn_plus = Button()
        btn_plus.Content = u"+"
        btn_plus.Style   = self.window.Resources["BtnSpin"]
        btn_plus.Tag     = "{0}:{1}:inc".format(group_id, side)
        btn_plus.Click  += self._on_spin_click

        container.Children.Add(btn_minus)
        container.Children.Add(lbl)
        container.Children.Add(btn_plus)
        return container

    def _build_segments_panel(self):
        panel = self._segments_panel
        panel.Children.Clear()
        self._diameter_combos.clear()

        segs = self.vm.splice_segments
        for seg in segs:
            row = Grid()
            row.Margin = Thickness(0, 2, 0, 2)

            for w in [55, -1, 75]:
                cd = ColumnDefinition()
                if w == -1:
                    cd.Width = _star_width()
                else:
                    cd.Width = _pixel_width(w)
                row.ColumnDefinitions.Add(cd)

            # Badge segmento coloreado
            badge = Border()
            badge.Background    = _hex_to_brush(seg.color)
            badge.CornerRadius  = _corner_radius(3)
            badge.Padding       = Thickness(6, 2, 6, 2)
            badge.HorizontalAlignment = HAlign.Center
            badge.VerticalAlignment   = VAlign.Center
            badge_lbl = TextBlock()
            badge_lbl.Text       = str(seg.segment_id)
            badge_lbl.FontSize   = 10
            badge_lbl.FontWeight = FW.Bold
            badge_lbl.Foreground = _hex_to_brush("#E8F4F8")
            badge_lbl.HorizontalAlignment = HAlign.Center
            badge.Child = badge_lbl
            _set_col(badge, 0)
            row.Children.Add(badge)

            # Rango de altura
            rng_lbl = TextBlock()
            rng_lbl.Text     = seg.level_range_label()
            rng_lbl.FontSize = 10
            rng_lbl.Foreground = _hex_to_brush("#95B8CC")
            rng_lbl.VerticalAlignment = VAlign.Center
            rng_lbl.Margin   = Thickness(4, 0, 0, 0)
            _set_col(rng_lbl, 1)
            row.Children.Add(rng_lbl)

            # ComboBox diámetro
            cb = ComboBox()
            try:
                cb.Style = self.window.Resources["ComboStyle"]
            except Exception:
                pass
            cb.Margin = Thickness(2, 0, 0, 0)
            sel_idx = 4  # Ø20 por defecto
            for idx_d, d in enumerate(self.AVAILABLE_DIAMETERS):
                item = ComboBoxItem()
                item.Content = u"Ø {0}".format(d)
                item.Tag     = d
                cb.Items.Add(item)
                if d == seg.diameter_mm:
                    sel_idx = idx_d
            cb.SelectedIndex = sel_idx
            cb.Tag = seg.segment_id
            cb.SelectionChanged += self._on_diameter_changed
            self._diameter_combos[seg.segment_id] = cb
            _set_col(cb, 2)
            row.Children.Add(cb)

            panel.Children.Add(row)

            sep = WpfRect()
            sep.Height = 1
            sep.Fill   = _hex_to_brush("#1A3040")
            sep.Margin = Thickness(0, 1, 0, 1)
            panel.Children.Add(sep)

        self._draw_segment_preview()

    # ------------------------------------------------------------------ #
    #  Lógica de paso activo (step indicator)                             #
    # ------------------------------------------------------------------ #

    def _set_active_step(self, step):
        self._current_step = step
        ACTIVE_BG    = "#1B6CA8"
        ACTIVE_BORDER = "#1B6CA8"
        ACTIVE_FG    = "#1B6CA8"
        DONE_BG      = "#0D3D5A"
        DONE_BORDER  = "#1B6CA8"
        DONE_FG      = "#E8F4F8"
        INACTIVE_BG  = "#0D2D45"
        INACTIVE_BD  = "#21465C"
        INACTIVE_FG  = "#4A7A94"

        for i in range(4):
            s = i + 1
            circle = self._step_circles[i]
            label  = self._step_labels[i]
            if s < step:  # completado
                circle.Background   = _hex_to_brush(DONE_BG)
                circle.BorderBrush  = _hex_to_brush(DONE_BORDER)
                txt = circle.Child
                txt.Foreground = _hex_to_brush(DONE_FG)
                txt.Text       = u"✓"
                label.Foreground = _hex_to_brush("#95B8CC")
            elif s == step:  # activo
                circle.Background  = _hex_to_brush(ACTIVE_BG)
                circle.BorderBrush = _hex_to_brush(ACTIVE_BORDER)
                txt = circle.Child
                txt.Foreground = _hex_to_brush("#E8F4F8")
                txt.Text       = str(s)
                label.Foreground = _hex_to_brush(ACTIVE_FG)
            else:  # inactivo
                circle.Background  = _hex_to_brush(INACTIVE_BG)
                circle.BorderBrush = _hex_to_brush(INACTIVE_BD)
                txt = circle.Child
                txt.Foreground = _hex_to_brush(INACTIVE_FG)
                txt.Text       = str(s)
                label.Foreground = _hex_to_brush(INACTIVE_FG)

        # Resaltar borde del panel activo
        for i, border in enumerate(self._panel_borders):
            if border is None:
                continue
            if i + 1 == step:
                border.BorderBrush = _hex_to_brush("#1B6CA8")
            else:
                border.BorderBrush = _hex_to_brush("#21465C")

        # Habilitar botón Generar solo en step 4
        gen_btn = self._find("BtnGenerate")
        if gen_btn is not None:
            gen_btn.IsEnabled = (step == 4)

    # ------------------------------------------------------------------ #
    #  Handlers de eventos                                                #
    # ------------------------------------------------------------------ #

    def _on_next1(self, sender, args):
        from column_reinforcement_v2.domain.validation_engine import ValidationEngine
        errors = ValidationEngine().validate_step1(self.vm.column_groups)
        if errors:
            self._show_validation_error(u"\n".join(errors))
            return
        self._find("BtnNext2").IsEnabled = True
        self._set_active_step(2)

    def _on_next2(self, sender, args):
        from column_reinforcement_v2.domain.validation_engine import ValidationEngine
        errors = ValidationEngine().validate_step2(self.vm.splice_segments)
        if errors:
            self._show_validation_error(u"\n".join(errors))
            return
        self._find("BtnNext3").IsEnabled = True
        self._set_active_step(3)

    def _on_next3(self, sender, args):
        from column_reinforcement_v2.domain.validation_engine import ValidationEngine
        errors = ValidationEngine().validate_step3(
            self.vm.column_groups, self.vm.distributions
        )
        if errors:
            self._show_validation_error(u"\n".join(errors))
            return
        self._find("BtnGenerate").IsEnabled = True
        self._set_active_step(4)

    def _on_back(self, sender, args):
        if self._current_step > 1:
            self._set_active_step(self._current_step - 1)

    def _on_generate(self, sender, args):
        from column_reinforcement_v2.domain.validation_engine import ValidationEngine
        errors = ValidationEngine().validate_step4(self.vm.splice_segments)
        if errors:
            self._show_validation_error(u"\n".join(errors))
            return
        self.request = self.vm.to_request()
        self.window.DialogResult = True
        self.window.Close()

    def _on_close(self, sender, args):
        self.request = None
        self.window.DialogResult = False
        self.window.Close()

    def _on_help(self, sender, args):
        if TaskDialog is not None:
            TaskDialog.Show(
                u"Arainco: Armado Columnas Wizard",
                u"Flujo de 4 pasos:\n"
                u"1. Selecciona las columnas agrupadas por sección.\n"
                u"2. Define los puntos de troceo/empalme.\n"
                u"3. Configura la distribución de barras por cara.\n"
                u"4. Asigna el diámetro por segmento.\n"
                u"Presiona 'Generar armado' para crear las barras en Revit.",
            )

    def _on_closed(self, sender, args):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

    def _on_splice_check_changed(self, sender, args):
        z_mm     = sender.Tag
        is_active = bool(sender.IsChecked)
        self.vm.set_cut_active(z_mm, is_active)
        self._refresh_step2()
        self._build_segments_panel()
        self._refresh_footer()

    def _on_spin_click(self, sender, args):
        tag_parts = str(sender.Tag).split(":")
        if len(tag_parts) < 3:
            return
        group_id = int(tag_parts[0])
        side     = tag_parts[1]
        action   = tag_parts[2]
        if action == "inc":
            self.vm.increment_bars(group_id, side)
        else:
            self.vm.decrement_bars(group_id, side)
        # Actualizar label del spinner
        dist = self.vm.distribution_for(group_id)
        if dist is not None:
            key = (group_id, side)
            if key in self._bar_labels:
                val = dist.side_a_count if side == "A" else dist.side_b_count
                self._bar_labels[key].Text = str(val)

    def _on_diameter_changed(self, sender, args):
        try:
            seg_id = int(sender.Tag)
        except Exception:
            return
        item = sender.SelectedItem
        if item is None:
            return
        try:
            # Tag puede ser Python int o CLR Int32 — convertir explícitamente
            d = int(item.Tag)
        except Exception:
            # Fallback: extraer número del texto "Ø 20"
            try:
                d = int(str(item.Content).replace(u"Ø ", u"").strip())
            except Exception:
                return
        self.vm.set_diameter(seg_id, d)
        self._draw_segment_preview()

    def _show_validation_error(self, message):
        if TaskDialog is not None:
            TaskDialog.Show(u"Arainco: Armado Columnas Wizard", message)

    # ------------------------------------------------------------------ #
    #  Mostrar                                                            #
    # ------------------------------------------------------------------ #

    def show_dialog(self):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, self.window)
        except Exception:
            pass
        result = self.window.ShowDialog()
        if result:
            return self.request
        return None


# ──────────────────────────────────────────────────────────────────────── #
#  Guard singleton                                                         #
# ──────────────────────────────────────────────────────────────────────── #

def show_singleton_wizard(column_elements):
    """Abre el wizard o enfoca la instancia existente."""
    try:
        existing = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
        if existing is not None:
            try:
                existing.Activate()
                existing.Focus()
            except Exception:
                pass
            if TaskDialog is not None:
                TaskDialog.Show(
                    u"Arainco: Armado Columnas Wizard",
                    u"La herramienta ya esta en ejecucion.",
                )
            return None
    except Exception:
        pass

    controller = WizardWindowController()
    controller.populate(column_elements)
    return controller.show_dialog()


# ──────────────────────────────────────────────────────────────────────── #
#  Helpers WPF (usan imports de módulo)                                    #
# ──────────────────────────────────────────────────────────────────────── #

def _fw_bold():
    return FW.Bold

def _fw_semibold():
    return FW.SemiBold

def _star_width():
    return GridLength(1, GridUnitType.Star)

def _pixel_width(px):
    return GridLength(px)

def _corner_radius(r):
    return CornerRadius(r)

def _set_col(elem, col):
    Grid.SetColumn(elem, col)
