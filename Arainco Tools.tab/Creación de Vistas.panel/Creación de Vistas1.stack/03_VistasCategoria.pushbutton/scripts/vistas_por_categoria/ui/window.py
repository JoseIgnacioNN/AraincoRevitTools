# -*- coding: utf-8 -*-
"""UI WPF — Vistas por Categoría."""

from __future__ import print_function

import weakref

try:
    unicode
except NameError:
    unicode = str

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import EventHandler
from System.Windows import FontWeights, RoutedEventHandler, Thickness
from System.Windows.Controls import (
    Button,
    CheckBox,
    ComboBoxItem,
    SelectionChangedEventHandler,
    TextChangedEventHandler,
)
from System.Windows.Input import Cursors
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes, SolidColorBrush, Color
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, TaskDialog
from Autodesk.Revit.DB import FilteredElementCollector, Level

from infra.bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from infra.revit_wpf_window_position import revit_main_hwnd

from vistas_por_categoria.constants import (
    CATEGORIA_OPTIONS,
    TRANSACTION_TITLE,
    VIEW_SCALE_RATIOS,
    ZONA_DEFAULT,
)
from vistas_por_categoria import singleton

_DIALOG_TITLE = TRANSACTION_TITLE
_WINDOW_TITLE = u"Arainco: Vistas por categoría"

_STATUS_IDLE = u"idle"
_STATUS_BUSY = u"busy"
_STATUS_OK = u"ok"
_STATUS_ERR = u"err"

_BRUSH_SEG_ON_BG = SolidColorBrush(Color.FromArgb(0x24, 0x5B, 0xB8, 0xD4))
_BRUSH_SEG_ON_BD = SolidColorBrush(Color.FromRgb(0x5B, 0xB8, 0xD4))
_BRUSH_SEG_OFF_BG = SolidColorBrush(Color.FromRgb(0x07, 0x10, 0x18))
_BRUSH_SEG_OFF_BD = SolidColorBrush(Color.FromRgb(0x1E, 0x33, 0x44))
_BRUSH_FG_HI = SolidColorBrush(Color.FromRgb(0xE8, 0xF4, 0xF8))
_BRUSH_FG_MID = SolidColorBrush(Color.FromRgb(0x95, 0xB8, 0xCC))
_BRUSH_FG_LO = SolidColorBrush(Color.FromRgb(0x64, 0x74, 0x8B))
_BRUSH_INFO = SolidColorBrush(Color.FromRgb(0x38, 0xBD, 0xF8))
_BRUSH_WARN = SolidColorBrush(Color.FromRgb(0xFB, 0xBF, 0x24))
_BRUSH_OK = SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80))
_BRUSH_ERR = SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71))
_BRUSH_STATUS_INFO_BG = SolidColorBrush(Color.FromArgb(0x14, 0x38, 0xBD, 0xF8))
_BRUSH_STATUS_BUSY_BG = SolidColorBrush(Color.FromArgb(0x1A, 0xFB, 0xBF, 0x24))
_BRUSH_STATUS_OK_BG = SolidColorBrush(Color.FromArgb(0x14, 0x4A, 0xDE, 0x80))
_BRUSH_STATUS_ERR_BG = SolidColorBrush(Color.FromArgb(0x14, 0xF8, 0x71, 0x71))
_BRUSH_STATUS_INFO_BD = SolidColorBrush(Color.FromArgb(0x4D, 0x38, 0xBD, 0xF8))
_BRUSH_STATUS_BUSY_BD = SolidColorBrush(Color.FromArgb(0x59, 0xFB, 0xBF, 0x24))
_BRUSH_STATUS_OK_BD = SolidColorBrush(Color.FromArgb(0x4D, 0x4A, 0xDE, 0x80))
_BRUSH_STATUS_ERR_BD = SolidColorBrush(Color.FromArgb(0x4D, 0xF8, 0x71, 0x71))
_BRUSH_TRANSPARENT = Brushes.Transparent


def _collect_levels_sorted(doc):
    levels = list(FilteredElementCollector(doc).OfClass(Level))
    levels.sort(key=lambda lv: lv.Elevation)
    return levels


def mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    hwnd = None
    try:
        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _DIALOG_TITLE,
            instruction,
            content=content,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(_DIALOG_TITLE, body)
    except Exception:
        pass


XAML = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco"
    Width="540"
    SizeToContent="Height"
    MinWidth="500"
    MaxHeight="860"
    WindowStartupLocation="Manual"
    Background="#071018"
    FontFamily="Segoe UI"
    FontSize="12"
    ShowInTaskbar="False">
  <Window.Resources>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
    <Style x:Key="Lbl" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#95B8CC"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
    </Style>
  </Window.Resources>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,12">
        <TextBlock Text="__TITLE__" FontSize="18" FontWeight="Bold" Foreground="#E8F4F8"/>
        <TextBlock Text="Plantas Cielo/Piso por nivel · plantillas y tipos Detail/Sección · clasificación 01_ENTREGABLE"
                   Foreground="#95B8CC" FontSize="11" Margin="0,6,0,0" TextWrapping="Wrap"/>
      </StackPanel>

      <StackPanel Grid.Row="1">
        <Border Background="#0a1620" BorderBrush="#21465C" BorderThickness="1"
                CornerRadius="4" Padding="12" Margin="0,0,0,10">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="1.55*"/>
              <ColumnDefinition Width="10"/>
              <ColumnDefinition Width="0.9*"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
              <TextBlock Text="Categoría de vistas" Style="{StaticResource Lbl}"/>
              <ComboBox x:Name="CmbCategoria" Style="{StaticResource Combo}" IsEditable="False"/>
              <Border Margin="0,8,0,0" Padding="6,4"
                      Background="#245BB8D4" BorderBrush="#595BB8D4" BorderThickness="1"
                      CornerRadius="4" HorizontalAlignment="Left">
                <StackPanel Orientation="Horizontal">
                  <TextBlock Text="Código" Foreground="#64748b" FontSize="10"
                             VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <TextBlock x:Name="TxtCodigo" Text="—" Foreground="#E8F4F8" FontSize="11"
                             FontFamily="Consolas" VerticalAlignment="Center"/>
                </StackPanel>
              </Border>
              <TextBlock Text="Nombre de la zona" Style="{StaticResource Lbl}" Margin="0,12,0,4"/>
              <TextBox x:Name="TxtZona" Style="{StaticResource BimToolsTextBoxDark}"
                       Text="GENERAL"/>
              <TextBlock Text="Si el proyecto no está dividido en zonas, use GENERAL."
                         Foreground="#64748b" FontSize="10" Margin="0,6,0,0" TextWrapping="Wrap"/>
            </StackPanel>
            <StackPanel Grid.Column="2">
              <TextBlock Text="Escala" Style="{StaticResource Lbl}"/>
              <WrapPanel x:Name="PanelEscala" Orientation="Horizontal"/>
            </StackPanel>
          </Grid>
        </Border>

        <Border Background="#0a1620" BorderBrush="#21465C" BorderThickness="1"
                CornerRadius="4" Padding="12" Margin="0,0,0,10">
          <StackPanel>
            <Grid Margin="0,0,0,6">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="Niveles" Style="{StaticResource Lbl}" Margin="0,0,8,0"
                           VerticalAlignment="Center"/>
                <Border Background="#245BB8D4" BorderBrush="#4D5BB8D4" BorderThickness="1"
                        CornerRadius="10" Padding="7,2" VerticalAlignment="Center">
                  <TextBlock x:Name="TxtNivelCount" Text="0 / 0" Foreground="#5BB8D4"
                             FontSize="10" FontWeight="SemiBold"/>
                </Border>
              </StackPanel>
              <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button x:Name="BtnSelAll" Content="Todos" Style="{StaticResource BtnSelectOutline}"
                        Padding="10,4" Margin="0,0,6,0"/>
                <Button x:Name="BtnSelNone" Content="Ninguno" Style="{StaticResource BtnSelectOutline}"
                        Padding="10,4"/>
              </StackPanel>
            </Grid>
            <Border BorderBrush="#1e3344" BorderThickness="1" CornerRadius="4"
                    Background="#071018" Padding="4,2">
              <ScrollViewer MaxHeight="220" VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="PanelNiveles"/>
              </ScrollViewer>
            </Border>
            <TextBlock x:Name="TxtNivelesHint" Text=""
                       Foreground="#64748b" FontSize="10" Margin="0,6,0,0" TextWrapping="Wrap"/>
          </StackPanel>
        </Border>
      </StackPanel>

      <StackPanel Grid.Row="2" Margin="0,2,0,0">
        <Border x:Name="BorderResumen" Background="#145BB8D4" BorderBrush="#595BB8D4"
                BorderThickness="1" CornerRadius="4" Padding="10,8" Margin="0,0,0,10">
          <TextBlock x:Name="TxtResumen" Foreground="#95B8CC" FontSize="11"
                     TextWrapping="Wrap"/>
        </Border>
        <Border x:Name="BorderEstado" Background="Transparent" BorderBrush="Transparent"
                BorderThickness="1" CornerRadius="4" Padding="0,0,0,0" Margin="0,0,0,10">
          <TextBlock x:Name="TxtEstado" Text="Listo para crear el conjunto 01_ENTREGABLE."
                     Foreground="#64748b" FontSize="11" TextWrapping="Wrap"/>
        </Border>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="10"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Button x:Name="BtnCancelar" Grid.Column="0" Content="Cancelar"
                  Style="{StaticResource BtnSelectOutline}"
                  Padding="16,9" MinWidth="110"/>
          <Button x:Name="BtnIniciar" Grid.Column="2" Content="Crear vistas"
                  Style="{StaticResource BtnPrimary}" HorizontalAlignment="Stretch"
                  Padding="12,9"/>
        </Grid>
      </StackPanel>
    </Grid>
  </Border>
</Window>
""".replace(u"__TITLE__", _WINDOW_TITLE)


class _LevelCheck(object):
    def __init__(self, level, checkbox):
        self.level = level
        self.checkbox = checkbox


class _CreateCategoriaViewsHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref
        self.request = None

    def Execute(self, uiapp):
        win = self._window_ref()
        req = self.request
        self.request = None
        if win is None or req is None:
            return

        from vistas_por_categoria.service import (
            VistasPorCategoriaError,
            create_categoria_views,
            format_success_dialog,
            validate_categoria_views_not_exist,
        )

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            mostrar_aviso(uiapp, u"No hay documento activo.")
            win._finish_create(False, u"No hay documento activo.")
            return

        doc = uidoc.Document
        try:
            ok, msg = validate_categoria_views_not_exist(
                doc, req.categoria_code, req.zona
            )
            if not ok:
                mostrar_aviso(uiapp, msg)
                win._finish_create(False, msg)
                return

            result = create_categoria_views(doc, req)
            instruction, content = format_success_dialog(
                result, req.categoria_display, req.categoria_code, req.zona
            )
            mostrar_aviso(uiapp, instruction, content, ok_text=u"Entendido")
            win._finish_create(
                True,
                u"Completado: {} vista(s) creada(s).".format(len(result.created)),
            )
        except VistasPorCategoriaError as ex:
            mostrar_aviso(uiapp, str(ex))
            win._finish_create(False, str(ex))
        except Exception as ex:
            mostrar_aviso(
                uiapp,
                u"Error al crear vistas.",
                content=u"{}".format(ex),
            )
            win._finish_create(False, u"Error: {}".format(ex))

    def GetName(self):
        return TRANSACTION_TITLE


class VistasPorCategoriaWindow(object):
    def __init__(self, doc, uidoc, revit_app):
        self._doc = doc
        self._uidoc = uidoc
        self._revit = revit_app
        self._level_checks = []
        self._scale_buttons = []
        self._scale_ratio = 100
        self._busy = False

        self._create_handler = _CreateCategoriaViewsHandler(weakref.ref(self))
        self._create_event = ExternalEvent.Create(self._create_handler)

        self._win = XamlReader.Parse(XAML)
        self._cmb_categoria = self._win.FindName("CmbCategoria")
        self._txt_zona = self._win.FindName("TxtZona")
        self._panel_escala = self._win.FindName("PanelEscala")
        self._txt_codigo = self._win.FindName("TxtCodigo")
        self._panel_niveles = self._win.FindName("PanelNiveles")
        self._txt_nivel_count = self._win.FindName("TxtNivelCount")
        self._txt_niveles_hint = self._win.FindName("TxtNivelesHint")
        self._txt_resumen = self._win.FindName("TxtResumen")
        self._border_estado = self._win.FindName("BorderEstado")
        self._txt_estado = self._win.FindName("TxtEstado")
        self._btn_iniciar = self._win.FindName("BtnIniciar")
        self._btn_cancelar = self._win.FindName("BtnCancelar")
        btn_all = self._win.FindName("BtnSelAll")
        btn_none = self._win.FindName("BtnSelNone")

        self._fill_categorias()
        self._fill_escalas()
        self._fill_niveles()
        if self._txt_zona is not None:
            self._txt_zona.Text = ZONA_DEFAULT
            try:
                self._txt_zona.TextChanged += TextChangedEventHandler(
                    self._on_zona_changed
                )
            except Exception:
                pass
        self._refresh_form_state()

        self._cmb_categoria.SelectionChanged += SelectionChangedEventHandler(
            self._on_categoria_changed
        )
        self._btn_iniciar.Click += RoutedEventHandler(self._on_iniciar)
        self._btn_cancelar.Click += RoutedEventHandler(lambda s, e: self._win.Close())
        if btn_all is not None:
            btn_all.Click += RoutedEventHandler(
                lambda s, e: self._set_all_levels(True)
            )
        if btn_none is not None:
            btn_none.Click += RoutedEventHandler(
                lambda s, e: self._set_all_levels(False)
            )

        self._win.Closed += EventHandler(lambda s, e: singleton.clear())

    def _on_categoria_changed(self, _sender, _e):
        self._refresh_form_state()

    def _on_zona_changed(self, _sender, _e):
        self._refresh_form_state()

    def _fill_categorias(self):
        self._cmb_categoria.Items.Clear()
        for code, label in CATEGORIA_OPTIONS:
            it = ComboBoxItem()
            it.Content = label
            it.Tag = code
            self._cmb_categoria.Items.Add(it)
        if self._cmb_categoria.Items.Count > 0:
            self._cmb_categoria.SelectedIndex = 0

    def _fill_escalas(self):
        self._panel_escala.Children.Clear()
        self._scale_buttons = []
        self._scale_ratio = 100
        for ratio in VIEW_SCALE_RATIOS:
            btn = Button()
            btn.Content = u"1:{}".format(ratio)
            btn.Tag = ratio
            btn.Margin = Thickness(0, 0, 4, 4)
            btn.Padding = Thickness(6, 6, 6, 6)
            btn.MinWidth = 52
            btn.FontSize = 12
            btn.Cursor = Cursors.Hand
            try:
                btn.Style = self._win.FindResource("BtnSelectOutline")
            except Exception:
                pass
            btn.Click += RoutedEventHandler(self._on_scale_click)
            self._panel_escala.Children.Add(btn)
            self._scale_buttons.append(btn)
        self._apply_scale_button_styles()

    def _on_scale_click(self, sender, _e):
        try:
            self._scale_ratio = int(sender.Tag)
        except Exception:
            self._scale_ratio = 100
        self._apply_scale_button_styles()
        self._refresh_form_state()

    def _apply_scale_button_styles(self):
        for btn in self._scale_buttons:
            on = False
            try:
                on = int(btn.Tag) == int(self._scale_ratio)
            except Exception:
                on = False
            if on:
                btn.Background = _BRUSH_SEG_ON_BG
                btn.BorderBrush = _BRUSH_SEG_ON_BD
                btn.Foreground = _BRUSH_FG_HI
                btn.FontWeight = FontWeights.SemiBold
            else:
                btn.Background = _BRUSH_SEG_OFF_BG
                btn.BorderBrush = _BRUSH_SEG_OFF_BD
                btn.Foreground = _BRUSH_FG_MID
                btn.FontWeight = FontWeights.Normal
            btn.BorderThickness = Thickness(1)

    def _fill_niveles(self):
        self._panel_niveles.Children.Clear()
        self._level_checks = []
        levels = _collect_levels_sorted(self._doc)
        for lv in levels:
            try:
                name = str(lv.Name or u"")
            except Exception:
                name = u"?"
            cb = CheckBox()
            cb.Content = name
            cb.IsChecked = True
            cb.Foreground = Brushes.White
            cb.Margin = Thickness(0, 4, 0, 4)
            cb.Tag = lv
            cb.Checked += RoutedEventHandler(self._on_level_changed)
            cb.Unchecked += RoutedEventHandler(self._on_level_changed)
            self._panel_niveles.Children.Add(cb)
            self._level_checks.append(_LevelCheck(lv, cb))

    def _on_level_changed(self, _sender, _e):
        self._refresh_form_state()

    def _set_all_levels(self, checked):
        for item in self._level_checks:
            item.checkbox.IsChecked = checked
        self._refresh_form_state()

    def _get_selected_categoria(self):
        sel = self._cmb_categoria.SelectedItem
        if sel is None:
            return None, None
        code = getattr(sel, "Tag", None)
        try:
            code = unicode(code).strip() if code is not None else u""
        except Exception:
            code = u""
        display = u""
        try:
            display = unicode(sel.Content)
        except Exception:
            display = code
        if not code and display:
            # Respaldo: parsear «08_RP - DETALLE…» como en Dynamo String.Split
            parts = display.split(u" - ", 1)
            if parts:
                code = parts[0].strip()
        return code, display

    def _get_zona(self):
        if self._txt_zona is None:
            return ZONA_DEFAULT
        try:
            z = unicode(self._txt_zona.Text or u"").strip()
        except Exception:
            z = u""
        return z or ZONA_DEFAULT

    def _get_selected_scale(self):
        try:
            return int(self._scale_ratio)
        except Exception:
            return 100

    def _get_selected_levels(self):
        out = []
        for item in self._level_checks:
            try:
                if item.checkbox.IsChecked:
                    out.append(item.level)
            except Exception:
                continue
        return out

    def _set_status(self, kind, text):
        if self._txt_estado is None:
            return
        self._txt_estado.Text = unicode(text or u"")
        if self._border_estado is None:
            return
        if kind == _STATUS_BUSY:
            self._border_estado.Background = _BRUSH_STATUS_BUSY_BG
            self._border_estado.BorderBrush = _BRUSH_STATUS_BUSY_BD
            self._border_estado.Padding = Thickness(10, 8, 10, 8)
            self._txt_estado.Foreground = _BRUSH_WARN
        elif kind == _STATUS_OK:
            self._border_estado.Background = _BRUSH_STATUS_OK_BG
            self._border_estado.BorderBrush = _BRUSH_STATUS_OK_BD
            self._border_estado.Padding = Thickness(10, 8, 10, 8)
            self._txt_estado.Foreground = _BRUSH_OK
        elif kind == _STATUS_ERR:
            self._border_estado.Background = _BRUSH_STATUS_ERR_BG
            self._border_estado.BorderBrush = _BRUSH_STATUS_ERR_BD
            self._border_estado.Padding = Thickness(10, 8, 10, 8)
            self._txt_estado.Foreground = _BRUSH_ERR
        elif kind == _STATUS_IDLE and text and not text.startswith(u"Listo"):
            self._border_estado.Background = _BRUSH_STATUS_INFO_BG
            self._border_estado.BorderBrush = _BRUSH_STATUS_INFO_BD
            self._border_estado.Padding = Thickness(10, 8, 10, 8)
            self._txt_estado.Foreground = _BRUSH_INFO
        else:
            self._border_estado.Background = _BRUSH_TRANSPARENT
            self._border_estado.BorderBrush = _BRUSH_TRANSPARENT
            self._border_estado.Padding = Thickness(0)
            self._txt_estado.Foreground = _BRUSH_FG_LO

    def _refresh_form_state(self):
        if self._busy:
            return

        code, display = self._get_selected_categoria()
        zona = self._get_zona()
        if self._txt_codigo is not None:
            self._txt_codigo.Text = unicode(code or u"—")

        if self._txt_niveles_hint is not None:
            self._txt_niveles_hint.Text = u""

        levels = self._get_selected_levels()
        total = len(self._level_checks)
        n = len(levels)
        if self._txt_nivel_count is not None:
            self._txt_nivel_count.Text = u"{0} / {1}".format(n, total)

        scale = self._get_selected_scale()
        if self._txt_resumen is not None:
            if not code:
                self._txt_resumen.Text = (
                    u"Seleccione una categoría para ver el resumen de creación."
                )
            elif n == 0:
                self._txt_resumen.Text = (
                    u"Seleccione al menos un nivel. Se crearán plantas Cielo/Piso, "
                    u"plantillas 01_ENTREGABLE, tipos Detail/Sección y filtro de sección."
                )
            else:
                plantas = n * 2
                self._txt_resumen.Text = (
                    u"Para {0} / zona {1} a 1:{2}: {3} plantas (Cielo+Piso) · "
                    u"plantillas 01_ENTREGABLE · tipos Detail/Sección · filtro de sección."
                ).format(code, zona, scale, plantas)

        can_run = bool(code) and n > 0
        if self._btn_iniciar is not None:
            self._btn_iniciar.IsEnabled = can_run

        if not code:
            self._set_status(_STATUS_ERR, u"Seleccione una categoría.")
        elif n == 0:
            self._set_status(
                _STATUS_ERR,
                u"Seleccione al menos un nivel para continuar.",
            )
        else:
            self._set_status(
                _STATUS_IDLE,
                u"Listo para crear el conjunto 01_ENTREGABLE.",
            )

    def _finish_create(self, success, status_text):
        self._busy = False
        self._refresh_form_state()
        kind = _STATUS_OK if success else _STATUS_ERR
        self._set_status(kind, status_text)

    def _on_iniciar(self, sender, args):
        code, display = self._get_selected_categoria()
        if not code:
            self._set_status(_STATUS_ERR, u"Seleccione una categoría.")
            mostrar_aviso(self._revit, u"Seleccione una categoría.")
            return

        levels = self._get_selected_levels()
        if not levels:
            self._set_status(_STATUS_ERR, u"Seleccione al menos un nivel.")
            mostrar_aviso(self._revit, u"Seleccione al menos un nivel.")
            return

        from vistas_por_categoria.service import VistasPorCategoriaRequest

        scale = self._get_selected_scale()
        zona = self._get_zona()
        req = VistasPorCategoriaRequest(code, zona, scale, levels, display)

        self._busy = True
        if self._btn_iniciar is not None:
            self._btn_iniciar.IsEnabled = False
        self._set_status(_STATUS_BUSY, u"Creando vistas… no cierre Revit.")
        self._create_handler.request = req
        self._create_event.Raise()

    def show(self):
        try:
            from System.Windows.Interop import WindowInteropHelper
            from infra.revit_wpf_window_position import (
                position_wpf_window_top_left_at_active_view,
            )

            hwnd = revit_main_hwnd(self._revit)
            if hwnd:
                WindowInteropHelper(self._win).Owner = hwnd
            position_wpf_window_top_left_at_active_view(self._win, self._uidoc, hwnd)
        except Exception:
            pass
        singleton.register(self._win)
        self._win.Show()


def show_vistas_por_categoria_ui(revit_app):
    if singleton.try_activate_existing():
        mostrar_aviso(revit_app, u"La herramienta ya está en ejecución.")
        return
    try:
        uidoc = revit_app.ActiveUIDocument
        doc = uidoc.Document
    except Exception:
        mostrar_aviso(revit_app, u"No hay documento activo.")
        return
    w = VistasPorCategoriaWindow(doc, uidoc, revit_app)
    w.show()
