# -*- coding: utf-8 -*-
"""UI WPF — Vistas por Categoría (shell visual alineado a Vistas por Usuario)."""

from __future__ import print_function

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

from System import AppDomain, EventHandler
from System.Windows import (
    CornerRadius,
    FontWeights,
    HorizontalAlignment,
    RoutedEventHandler,
    Thickness,
    VerticalAlignment,
)
from System.Windows.Controls import (
    Border,
    Button,
    CheckBox,
    ComboBoxItem,
    Orientation,
    SelectionChangedEventHandler,
    StackPanel,
    TextBlock,
    TextChangedEventHandler,
)
from System.Windows.Input import Cursors, MouseButtonEventHandler
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes, Color, SolidColorBrush
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, TaskDialog
from Autodesk.Revit.DB import FilteredElementCollector, Level

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

from vistas_por_categoria.constants import (
    CATEGORIA_OPTIONS,
    TRANSACTION_TITLE,
    VIEW_SCALE_RATIOS,
    ZONA_DEFAULT,
)
from vistas_por_categoria import singleton

_DIALOG_TITLE = TRANSACTION_TITLE
_WINDOW_TITLE = u"Arainco: Vistas por categoría"
_DEFAULT_VIEW_SCALE = 50
_APPDOMAIN_EVENT_KEY = u"Arainco_VistasPorCategoria_ExtEvent"
_APPDOMAIN_HANDLER_KEY = u"Arainco_VistasPorCategoria_Handler"

_BRUSH_SEG_ON_BG = SolidColorBrush(Color.FromArgb(0x24, 0x5B, 0xB8, 0xD4))
_BRUSH_SEG_ON_BD = SolidColorBrush(Color.FromRgb(0x5B, 0xB8, 0xD4))
_BRUSH_SEG_OFF_BG = SolidColorBrush(Color.FromRgb(0x07, 0x10, 0x18))
_BRUSH_SEG_OFF_BD = SolidColorBrush(Color.FromRgb(0x1E, 0x33, 0x44))
_BRUSH_FG_HI = SolidColorBrush(Color.FromRgb(0xE8, 0xF4, 0xF8))
_BRUSH_FG_MID = SolidColorBrush(Color.FromRgb(0x95, 0xB8, 0xCC))
_BRUSH_ROW_HOVER = SolidColorBrush(Color.FromArgb(0x28, 0x5B, 0xC0, 0xDE))

_BODY_XAML = u"""
<StackPanel>
  <!-- 1. Categoría + zona -->
  <TextBlock Text="Categoría de vistas" Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             Margin="0,0,0,4"/>
  <ComboBox x:Name="CmbCategoria" Style="{StaticResource Combo}" IsEditable="False"
            Margin="0,0,0,8"/>
  <Border x:Name="ChipCodigo" Padding="7,3" Margin="0,0,0,14"
          Background="#245BB8D4" BorderBrush="#4D5BB8D4" BorderThickness="1"
          CornerRadius="10" HorizontalAlignment="Left" VerticalAlignment="Center">
    <StackPanel Orientation="Horizontal">
      <TextBlock Text="Código" Foreground="#64748b" FontSize="10"
                 VerticalAlignment="Center" Margin="0,0,8,0"/>
      <TextBlock x:Name="TxtCodigo" Text="—" Foreground="#E8F4F8" FontSize="11"
                 FontFamily="Consolas" FontWeight="SemiBold" VerticalAlignment="Center"/>
    </StackPanel>
  </Border>

  <TextBlock Text="Nombre de la zona" Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
             Margin="0,0,0,4"/>
  <TextBox x:Name="TxtZona" Style="{StaticResource BimToolsTextBoxDark}"
           Text="GENERAL" Margin="0,0,0,4"/>
  <TextBlock Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,0,14"
             Text="Si el proyecto no está dividido en zonas, use GENERAL."/>

  <!-- 2. Niveles -->
  <Grid Margin="0,0,0,8">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
      <TextBlock Text="Niveles" Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
                 VerticalAlignment="Center"/>
      <TextBlock x:Name="TxtNivelCount" Margin="10,0,0,0" VerticalAlignment="Center"
                 Foreground="#64748b" FontSize="11" Text="0 de 0 seleccionados"/>
    </StackPanel>
    <StackPanel Grid.Column="1" Orientation="Horizontal">
      <Button x:Name="BtnSelAll" Content="Seleccionar todo" Margin="0,0,6,0"
              Style="{StaticResource BtnSelectOutline}" MinWidth="120" Padding="8,2"
              ToolTip="Marcar todos los niveles"/>
      <Button x:Name="BtnSelNone" Content="Ninguno"
              Style="{StaticResource BtnSelectOutline}" MinWidth="72" Padding="8,2"
              ToolTip="Desmarcar todos los niveles"/>
    </StackPanel>
  </Grid>
  <Border Background="#050E18" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" MaxHeight="280">
    <ScrollViewer VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Disabled"
                  Padding="2,2">
      <StackPanel x:Name="PanelNiveles"/>
    </ScrollViewer>
  </Border>

  <!-- 3. Escala -->
  <TextBlock Margin="0,14,0,6" Text="Escala de vista"
             Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"/>
  <WrapPanel x:Name="PanelEscala" Orientation="Horizontal"/>
  <TextBlock Margin="0,8,0,0" Foreground="#64748b" FontSize="10" TextWrapping="Wrap"
             Text="También crea plantillas y tipos Detail/Sección de la categoría."/>
</StackPanel>
"""

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnCancelar" Content="Cerrar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="100"/>
<Button x:Name="BtnIniciar" Content="Crear vistas"
        Style="{StaticResource BtnPrimary}" MinWidth="150"
        ToolTip="Crear el conjunto 01_ENTREGABLE para la categoría y zona seleccionadas"/>
"""


def _collect_levels_sorted(doc):
    """Niveles del documento ordenados por elevación."""
    levels = list(FilteredElementCollector(doc).OfClass(Level))
    levels.sort(key=lambda lv: lv.Elevation)
    return levels


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    """Diálogo informativo WPF (estilo BIMTools). Respaldo: TaskDialog nativo."""
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


def _attach_revit_owner(win, uiapp):
    if win is None or uiapp is None:
        return
    try:
        from System.Windows.Interop import WindowInteropHelper

        hwnd = revit_main_hwnd(uiapp)
        if hwnd is not None:
            WindowInteropHelper(win).Owner = hwnd
    except Exception:
        pass


def _prepare_window(win, uiapp):
    if win is None:
        return
    try:
        hwnd = revit_main_hwnd(uiapp)
        bind_center_wpf_on_revit_monitor(win, hwnd)
        position_wpf_window_center_on_monitor(win, hwnd)
    except Exception:
        pass
    _attach_revit_owner(win, uiapp)


def _build_xaml():
    return build_simple_tool_xaml(
        title=_WINDOW_TITLE,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=_BODY_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        resize_mode=u"CanResizeWithGrip",
        size_to_content_height=True,
    )


def _pin_external_event(ext_event, handler):
    """Mantener vivos ExternalEvent/handler tras cerrar la UI WPF."""
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_EVENT_KEY, ext_event)
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_HANDLER_KEY, handler)
    except Exception:
        pass


def _unpin_external_event():
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_EVENT_KEY, None)
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_HANDLER_KEY, None)
    except Exception:
        pass


def _style_crear_button(btn, enabled):
    """CTA primario visible; deshabilitado se ve claramente apagado."""
    if btn is None:
        return
    try:
        btn.IsEnabled = bool(enabled)
    except Exception:
        pass
    try:
        if enabled:
            btn.Opacity = 1.0
            btn.Background = SolidColorBrush(Color.FromRgb(0x5B, 0xC0, 0xDE))
            btn.Foreground = SolidColorBrush(Color.FromRgb(0x0A, 0x1A, 0x2F))
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(0x87, 0xD9, 0xEE))
        else:
            btn.Opacity = 0.38
            btn.Background = SolidColorBrush(Color.FromRgb(0x1E, 0x33, 0x44))
            btn.Foreground = SolidColorBrush(Color.FromRgb(0x64, 0x74, 0x8B))
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(0x21, 0x46, 0x5C))
    except Exception:
        try:
            btn.Opacity = 1.0 if enabled else 0.4
        except Exception:
            pass


class _LevelCheck(object):
    def __init__(self, level, checkbox):
        self.level = level
        self.checkbox = checkbox


class _CreateCategoriaViewsHandler(IExternalEventHandler):
    """Ejecuta la creación en el hilo de Revit (UI ya cerrada)."""

    def __init__(self):
        self.request = None
        self.uiapp_for_dialog = None

    def Execute(self, uiapp):
        req = self.request
        self.request = None
        host = self.uiapp_for_dialog or uiapp
        if req is None:
            _unpin_external_event()
            return

        from vistas_por_categoria.service import (
            VistasPorCategoriaError,
            create_categoria_views,
            format_success_dialog,
            validate_categoria_views_not_exist,
        )

        try:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is None:
                mostrar_aviso(host, u"No hay documento activo.")
                return

            doc = uidoc.Document
            ok, msg = validate_categoria_views_not_exist(
                doc, req.categoria_code, req.zona
            )
            if not ok:
                mostrar_aviso(host, msg)
                return

            result = create_categoria_views(doc, req)
            instruction, content = format_success_dialog(
                result, req.categoria_display, req.categoria_code, req.zona
            )
            mostrar_aviso(host, instruction, content, ok_text=u"Entendido")
        except VistasPorCategoriaError as ex:
            mostrar_aviso(host, str(ex))
        except Exception as ex:
            mostrar_aviso(
                host,
                u"Error al crear vistas.",
                content=u"{}".format(ex),
            )
        finally:
            _unpin_external_event()

    def GetName(self):
        return TRANSACTION_TITLE


class VistasPorCategoriaWindow(object):
    def __init__(self, doc, uidoc, revit_app):
        self._doc = doc
        self._uidoc = uidoc
        self._revit = revit_app
        self._level_checks = []
        self._scale_buttons = []
        self._scale_ratio = int(_DEFAULT_VIEW_SCALE)
        self._busy = False

        self._create_handler = _CreateCategoriaViewsHandler()
        self._create_event = ExternalEvent.Create(self._create_handler)

        self._win = XamlReader.Parse(_build_xaml())
        self._cmb_categoria = self._win.FindName("CmbCategoria")
        self._txt_zona = self._win.FindName("TxtZona")
        self._panel_escala = self._win.FindName("PanelEscala")
        self._txt_codigo = self._win.FindName("TxtCodigo")
        self._panel_niveles = self._win.FindName("PanelNiveles")
        self._txt_nivel_count = self._win.FindName("TxtNivelCount")
        self._txt_subtitle = self._win.FindName("TxtSubtitle")
        self._txt_status = self._win.FindName("TxtStatus")
        self._btn_iniciar = self._win.FindName("BtnIniciar")
        self._btn_cancelar = self._win.FindName("BtnCancelar")
        btn_all = self._win.FindName("BtnSelAll")
        btn_none = self._win.FindName("BtnSelNone")

        if self._txt_subtitle is not None:
            try:
                self._txt_subtitle.Text = u"Plantas Cielo/Piso · 01_ENTREGABLE"
            except Exception:
                pass

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
        self._scale_ratio = int(_DEFAULT_VIEW_SCALE)
        for ratio in VIEW_SCALE_RATIOS:
            btn = Button()
            btn.Content = u"1:{}".format(ratio)
            btn.Tag = ratio
            btn.Margin = Thickness(0, 0, 4, 4)
            btn.Padding = Thickness(8, 6, 8, 6)
            btn.MinWidth = 52
            btn.FontSize = 12
            btn.Cursor = Cursors.Hand
            btn.Click += RoutedEventHandler(self._on_scale_click)
            self._panel_escala.Children.Add(btn)
            self._scale_buttons.append(btn)
        self._apply_scale_button_styles()

    def _on_scale_click(self, sender, _e):
        try:
            self._scale_ratio = int(sender.Tag)
        except Exception:
            self._scale_ratio = int(_DEFAULT_VIEW_SCALE)
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
                name = _as_unicode(lv.Name or u"")
            except Exception:
                name = u"?"

            cb = CheckBox()
            try:
                cb.Content = u""
            except Exception:
                cb.Content = None
            cb.IsChecked = True
            cb.Margin = Thickness(0, 0, 0, 0)
            cb.Padding = Thickness(0, 0, 0, 0)
            cb.Cursor = Cursors.Hand
            cb.VerticalAlignment = VerticalAlignment.Center
            cb.VerticalContentAlignment = VerticalAlignment.Center
            cb.HorizontalAlignment = HorizontalAlignment.Left
            cb.Tag = lv

            label = TextBlock()
            label.Text = name
            label.FontSize = 12
            label.Margin = Thickness(8, 0, 0, 0)
            label.Padding = Thickness(0, 0, 0, 0)
            label.VerticalAlignment = VerticalAlignment.Center
            label.Cursor = Cursors.Hand
            try:
                label.Foreground = _BRUSH_FG_HI
            except Exception:
                pass

            row_panel = StackPanel()
            row_panel.Orientation = Orientation.Horizontal
            row_panel.VerticalAlignment = VerticalAlignment.Center
            row_panel.Children.Add(cb)
            row_panel.Children.Add(label)

            border = Border()
            border.Background = Brushes.Transparent
            border.CornerRadius = CornerRadius(2)
            border.Padding = Thickness(6, 3, 6, 3)
            border.Margin = Thickness(0, 0, 0, 0)
            border.Cursor = Cursors.Hand
            border.Child = row_panel
            border.Tag = lv

            def _make_toggle(checkbox):
                def _on_row_down(sender, args):
                    try:
                        src = args.OriginalSource
                    except Exception:
                        src = None
                    if isinstance(src, CheckBox):
                        return
                    try:
                        checkbox.IsChecked = not bool(checkbox.IsChecked)
                        if args is not None:
                            args.Handled = True
                    except Exception:
                        pass

                return _on_row_down

            try:
                border.MouseLeftButtonDown += MouseButtonEventHandler(
                    _make_toggle(cb)
                )
            except Exception:
                pass

            def _make_hover(b):
                def _enter(sender, args):
                    try:
                        b.Background = _BRUSH_ROW_HOVER
                    except Exception:
                        pass

                def _leave(sender, args):
                    try:
                        b.Background = Brushes.Transparent
                    except Exception:
                        pass

                return _enter, _leave

            enter, leave = _make_hover(border)
            try:
                from System.Windows.Input import MouseEventHandler

                border.MouseEnter += MouseEventHandler(enter)
                border.MouseLeave += MouseEventHandler(leave)
            except Exception:
                pass

            cb.Checked += RoutedEventHandler(self._on_level_changed)
            cb.Unchecked += RoutedEventHandler(self._on_level_changed)
            self._panel_niveles.Children.Add(border)
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
            return int(_DEFAULT_VIEW_SCALE)

    def _get_selected_levels(self):
        out = []
        for item in self._level_checks:
            try:
                if item.checkbox.IsChecked:
                    out.append(item.level)
            except Exception:
                continue
        return out

    def _set_status(self, text):
        if self._txt_status is None:
            return
        try:
            self._txt_status.Text = _as_unicode(text or u"")
        except Exception:
            pass

    def _refresh_form_state(self):
        if self._busy:
            return

        code, display = self._get_selected_categoria()
        zona = self._get_zona()
        if self._txt_codigo is not None:
            self._txt_codigo.Text = unicode(code or u"—")

        levels = self._get_selected_levels()
        total = len(self._level_checks)
        n = len(levels)
        if self._txt_nivel_count is not None:
            if total <= 0:
                self._txt_nivel_count.Text = u"0 de 0 seleccionados"
            else:
                self._txt_nivel_count.Text = u"{0} de {1} seleccionados".format(
                    n, total
                )

        scale = self._get_selected_scale()
        can_run = bool(code) and n > 0
        _style_crear_button(self._btn_iniciar, can_run)

        if not code:
            self._set_status(u"Seleccione una categoría.")
        elif n == 0:
            self._set_status(u"Seleccione al menos un nivel.")
        else:
            self._set_status(
                u"{0} / zona {1} · 1:{2} · {3} plantas".format(
                    code, zona, scale, n * 2
                )
            )

    def _on_iniciar(self, sender, args):
        if self._busy:
            return

        code, display = self._get_selected_categoria()
        if not code:
            self._set_status(u"Seleccione una categoría.")
            mostrar_aviso(self._revit, u"Seleccione una categoría.")
            return

        levels = self._get_selected_levels()
        if not levels:
            self._set_status(u"Seleccione al menos un nivel.")
            mostrar_aviso(self._revit, u"Seleccione al menos un nivel.")
            return

        from vistas_por_categoria.service import VistasPorCategoriaRequest

        scale = self._get_selected_scale()
        zona = self._get_zona()
        req = VistasPorCategoriaRequest(code, zona, scale, list(levels), display)

        self._busy = True
        _style_crear_button(self._btn_iniciar, False)

        self._create_handler.request = req
        self._create_handler.uiapp_for_dialog = self._revit
        _pin_external_event(self._create_event, self._create_handler)
        try:
            self._create_event.Raise()
        except Exception as ex:
            self._busy = False
            _unpin_external_event()
            self._refresh_form_state()
            mostrar_aviso(
                self._revit,
                u"No se pudo iniciar la creación.",
                content=u"{}".format(ex),
            )
            return

        # Cerrar UI de inmediato; Execute usa el snapshot (sin controles WPF).
        try:
            self._win.Close()
        except Exception:
            pass

    def show(self):
        _prepare_window(self._win, self._revit)
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
