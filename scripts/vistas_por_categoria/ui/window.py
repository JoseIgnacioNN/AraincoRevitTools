# -*- coding: utf-8 -*-
"""UI WPF — Vistas por Categoría (selección múltiple de categorías, misma zona)."""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import os

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
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    RoutedEventHandler,
    Thickness,
    VerticalAlignment,
    Visibility,
)
from System.Windows.Controls import (
    Border,
    Button,
    CheckBox,
    ColumnDefinition,
    Grid,
    Orientation,
    StackPanel,
    TextBlock,
    TextChangedEventHandler,
)
from System.Windows.Input import Cursors, MouseButtonEventHandler, MouseEventHandler
from System.Windows.Markup import XamlReader
from System.Windows.Media import Brushes, Color, FontFamily, SolidColorBrush
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, TaskDialog
from Autodesk.Revit.DB import FilteredElementCollector, Level

from bimtools_ui_tokens import BTN_MANUAL
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
_FILTER_PLACEHOLDER = u"Filtrar por código o nombre…"
_SUBTITLE = u"01_ENTREGABLE · Cielo + Piso por nivel (LO: 4 plantas/nivel)"

_BRUSH_SEG_ON_BG = SolidColorBrush(Color.FromArgb(0x24, 0x5B, 0xB8, 0xD4))
_BRUSH_SEG_ON_BD = SolidColorBrush(Color.FromRgb(0x5B, 0xB8, 0xD4))
_BRUSH_SEG_OFF_BG = SolidColorBrush(Color.FromRgb(0x07, 0x10, 0x18))
_BRUSH_SEG_OFF_BD = SolidColorBrush(Color.FromRgb(0x1E, 0x33, 0x44))
_BRUSH_FG_HI = SolidColorBrush(Color.FromRgb(0xE8, 0xF4, 0xF8))
_BRUSH_FG_MID = SolidColorBrush(Color.FromRgb(0x95, 0xB8, 0xCC))
_BRUSH_FG_MUTED = SolidColorBrush(Color.FromRgb(0x64, 0x74, 0x8B))
_BRUSH_ROW_HOVER = SolidColorBrush(Color.FromArgb(0x28, 0x5B, 0xC0, 0xDE))

_SCROLLBAR_LOCAL = u"""
    <Border.Resources>
      <Style TargetType="ScrollBar" BasedOn="{StaticResource BimToolsScrollBarDark}"/>
    </Border.Resources>
"""

_BODY_XAML = u"""
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>

  <Grid Grid.Row="0" Margin="0,0,0,8">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
      <TextBlock Text="Categorías" Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
                 VerticalAlignment="Center"/>
      <TextBlock x:Name="TxtCatCount" Margin="10,0,0,0" VerticalAlignment="Center"
                 Foreground="#64748b" FontSize="11" Text="0 de 0 seleccionadas"/>
    </StackPanel>
    <StackPanel Grid.Column="1" Orientation="Horizontal">
      <Button x:Name="BtnCatSelAll" Content="Seleccionar todo" Margin="0,0,6,0"
              Style="{StaticResource BtnSelectOutline}" MinWidth="120" Padding="8,2"
              ToolTip="Marcar las categorías visibles del filtro"/>
      <Button x:Name="BtnCatSelNone" Content="Ninguna"
              Style="{StaticResource BtnSelectOutline}" MinWidth="72" Padding="8,2"
              ToolTip="Desmarcar las categorías visibles del filtro"/>
    </StackPanel>
  </Grid>

  <Grid Grid.Row="1" Margin="0,0,0,8" MinHeight="30">
    <TextBox x:Name="TxtFilter" MinHeight="30"
             Style="{StaticResource BimToolsTextBoxDark}"
             ToolTip="Filtrar categorías por código o nombre"
             VerticalContentAlignment="Center" Padding="10,4,10,4"/>
    <TextBlock x:Name="TxtFilterPh" Text="Filtrar por código o nombre…"
               IsHitTestVisible="False" Foreground="#64748b" FontSize="11"
               VerticalAlignment="Center" Margin="12,0,0,0"/>
  </Grid>

  <Border Grid.Row="2" Background="#050E18" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" MinHeight="160">
""" + _SCROLLBAR_LOCAL + u"""
    <ScrollViewer VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Disabled"
                  Padding="2,2">
      <StackPanel x:Name="PanelCategorias"/>
    </ScrollViewer>
  </Border>

  <TextBlock x:Name="TxtFilterEmpty" Grid.Row="3" Margin="0,6,0,0" Visibility="Collapsed"
             Foreground="#64748b" FontSize="10"
             Text="Ninguna categoría coincide con el filtro."/>

  <StackPanel Grid.Row="4" Margin="0,12,0,12">
    <Grid Margin="0,0,0,10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <TextBlock Text="Zona" Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
                 VerticalAlignment="Center" Margin="0,0,10,0"/>
      <TextBox x:Name="TxtZona" Grid.Column="1" Style="{StaticResource BimToolsTextBoxDark}"
               Text="GENERAL" MinHeight="30" VerticalContentAlignment="Center"
               ToolTip="Si el proyecto no está dividido en zonas, use GENERAL."/>
    </Grid>
    <TextBlock Text="Escala de vista" Foreground="#95B8CC" FontSize="11" FontWeight="SemiBold"
               Margin="0,0,0,6"/>
    <WrapPanel x:Name="PanelEscala" Orientation="Horizontal"/>
  </StackPanel>

  <Grid Grid.Row="5" Margin="0,0,0,8">
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

  <Border Grid.Row="6" Background="#050E18" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" MaxHeight="140">
""" + _SCROLLBAR_LOCAL + u"""
    <ScrollViewer VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Disabled"
                  Padding="2,2">
      <StackPanel x:Name="PanelNiveles"/>
    </ScrollViewer>
  </Border>
</Grid>
"""

_FOOTER_LEADING_XAML = (
    u'<Button x:Name="BtnManual" Content="Manual" '
    u'Style="{{StaticResource BtnSelectOutline}}" '
    u'Background="{bg}" MinWidth="96" Padding="8,2" '
    u'ToolTip="Abrir manual de usuario" VerticalAlignment="Center"/>'
).format(bg=BTN_MANUAL)

_FOOTER_ACTIONS_XAML = u"""
<Button x:Name="BtnCancelar" Content="Cerrar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="100"/>
<Button x:Name="BtnIniciar" Content="Crear vistas"
        Style="{StaticResource BtnPrimary}" MinWidth="150"
        ToolTip="Crear el conjunto 01_ENTREGABLE (Cielo y Piso por nivel). También crea plantillas y tipos Detail/Sección de cada categoría."/>
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


def _categoria_nombre(code, display):
    """Etiqueta sin el código duplicado («PLANTAS LOSAS» desde «01_LO - PLANTAS LOSAS»)."""
    d = _as_unicode(display).strip()
    c = _as_unicode(code).strip()
    prefix = c + u" - "
    if c and d.startswith(prefix):
        return d[len(prefix):].strip() or d
    sep = u" - "
    if sep in d:
        return d.split(sep, 1)[1].strip() or d
    return d


def _visibility_collapsed():
    return Visibility.Collapsed


def _visibility_visible():
    return Visibility.Visible


def _sync_filter_placeholder(txt, txt_ph):
    if txt_ph is None:
        return
    try:
        empty = not _as_unicode(txt.Text if txt is not None else u"").strip()
    except Exception:
        empty = True
    try:
        txt_ph.Visibility = (
            _visibility_visible() if empty else _visibility_collapsed()
        )
    except Exception:
        pass


def _filter_query(txt):
    if txt is None:
        return u""
    try:
        return _as_unicode(txt.Text).strip().lower()
    except Exception:
        return u""


def _make_check_row(tag, label_text, is_checked, on_changed, meta_text=None):
    """Fila clicable con checkbox + etiqueta (+ código opcional). Devuelve (border, checkbox)."""
    cb = CheckBox()
    try:
        cb.Content = u""
    except Exception:
        cb.Content = None
    cb.IsChecked = bool(is_checked)
    cb.Margin = Thickness(0, 0, 0, 0)
    cb.Padding = Thickness(0, 0, 0, 0)
    cb.Cursor = Cursors.Hand
    cb.VerticalAlignment = VerticalAlignment.Center
    cb.VerticalContentAlignment = VerticalAlignment.Center
    cb.HorizontalAlignment = HorizontalAlignment.Left
    cb.Tag = tag

    label = TextBlock()
    label.Text = _as_unicode(label_text)
    label.FontSize = 12
    label.Margin = Thickness(8, 0, 0, 0)
    label.Padding = Thickness(0, 0, 0, 0)
    label.VerticalAlignment = VerticalAlignment.Center
    label.Cursor = Cursors.Hand
    try:
        label.Foreground = _BRUSH_FG_HI
    except Exception:
        pass

    left = StackPanel()
    left.Orientation = Orientation.Horizontal
    left.VerticalAlignment = VerticalAlignment.Center
    left.Children.Add(cb)
    left.Children.Add(label)

    meta = _as_unicode(meta_text).strip() if meta_text else u""
    if meta:
        row_grid = Grid()
        col_star = ColumnDefinition()
        col_star.Width = GridLength(1, GridUnitType.Star)
        col_auto = ColumnDefinition()
        col_auto.Width = GridLength(1, GridUnitType.Auto)
        row_grid.ColumnDefinitions.Add(col_star)
        row_grid.ColumnDefinitions.Add(col_auto)
        left.VerticalAlignment = VerticalAlignment.Center
        left.HorizontalAlignment = HorizontalAlignment.Stretch
        Grid.SetColumn(left, 0)
        row_grid.Children.Add(left)
        meta_tb = TextBlock()
        meta_tb.Text = meta
        meta_tb.FontSize = 10
        meta_tb.Margin = Thickness(8, 0, 0, 0)
        meta_tb.VerticalAlignment = VerticalAlignment.Center
        meta_tb.Cursor = Cursors.Hand
        try:
            meta_tb.Foreground = _BRUSH_FG_MUTED
            meta_tb.FontFamily = FontFamily(u"Consolas")
        except Exception:
            pass
        Grid.SetColumn(meta_tb, 1)
        row_grid.Children.Add(meta_tb)
        child = row_grid
    else:
        child = left

    border = Border()
    border.Background = Brushes.Transparent
    border.CornerRadius = CornerRadius(2)
    border.Padding = Thickness(6, 3, 6, 3)
    border.Margin = Thickness(0, 0, 0, 0)
    border.Cursor = Cursors.Hand
    border.Child = child
    border.Tag = tag

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
        border.MouseLeftButtonDown += MouseButtonEventHandler(_make_toggle(cb))
    except Exception:
        pass

    holder = {u"cb": cb, u"border": border, u"hover": False}

    def _enter(sender, args):
        holder[u"hover"] = True
        try:
            if holder[u"cb"].IsChecked:
                holder[u"border"].Background = _BRUSH_SEG_ON_BG
            else:
                holder[u"border"].Background = _BRUSH_ROW_HOVER
        except Exception:
            pass

    def _leave(sender, args):
        holder[u"hover"] = False
        try:
            if holder[u"cb"].IsChecked:
                holder[u"border"].Background = _BRUSH_SEG_ON_BG
            else:
                holder[u"border"].Background = Brushes.Transparent
        except Exception:
            pass

    try:
        border.MouseEnter += MouseEventHandler(_enter)
        border.MouseLeave += MouseEventHandler(_leave)
    except Exception:
        pass

    def _on_check(sender, args):
        try:
            if holder[u"cb"].IsChecked:
                holder[u"border"].Background = _BRUSH_SEG_ON_BG
            elif holder[u"hover"]:
                holder[u"border"].Background = _BRUSH_ROW_HOVER
            else:
                holder[u"border"].Background = Brushes.Transparent
        except Exception:
            pass
        if on_changed is not None:
            on_changed(sender, args)

    cb.Checked += RoutedEventHandler(_on_check)
    cb.Unchecked += RoutedEventHandler(_on_check)

    try:
        if cb.IsChecked:
            border.Background = _BRUSH_SEG_ON_BG
    except Exception:
        pass

    return border, cb, holder


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


def _resolve_manual_path():
    """Ruta a ``manual_usuario.html`` en la carpeta del pushbutton."""
    candidates = []
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb:
            candidates.append(os.path.join(pb, u"manual_usuario.html"))
    except Exception:
        pass
    try:
        cursor = os.path.dirname(os.path.abspath(__file__))
        ext_dir = None
        for _ in range(16):
            try:
                names = os.listdir(cursor)
            except Exception:
                names = []
            if any(_as_unicode(n).endswith(u".tab") for n in names):
                ext_dir = cursor
                break
            parent = os.path.dirname(cursor)
            if parent == cursor:
                break
            cursor = parent
        if ext_dir:
            for tab_name in os.listdir(ext_dir):
                if not _as_unicode(tab_name).endswith(u".tab"):
                    continue
                panel = os.path.join(ext_dir, tab_name, u"Creación de Vistas.panel")
                if not os.path.isdir(panel):
                    panel = os.path.join(
                        ext_dir, tab_name, u"Creacion de Vistas.panel",
                    )
                if not os.path.isdir(panel):
                    continue
                for pb_name in os.listdir(panel):
                    if u"VistasCategoria" not in _as_unicode(pb_name):
                        continue
                    candidates.append(
                        os.path.join(panel, pb_name, u"manual_usuario.html")
                    )
    except Exception:
        pass
    seen = set()
    for path in candidates:
        try:
            ap = os.path.normpath(os.path.abspath(path))
        except Exception:
            continue
        if ap in seen:
            continue
        seen.add(ap)
        if os.path.isfile(ap):
            return ap
    return None


def _open_manual(uiapp):
    path = _resolve_manual_path()
    if not path:
        mostrar_aviso(
            uiapp,
            u"No se encontró manual_usuario.html.",
            content=u"Debe estar en la carpeta del pushbutton de la herramienta.",
        )
        return
    try:
        os.startfile(path)
    except Exception as ex:
        mostrar_aviso(
            uiapp,
            u"No se pudo abrir el manual.",
            content=_as_unicode(ex),
        )


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
        footer_leading_xaml=_FOOTER_LEADING_XAML,
        footer_actions_xaml=_FOOTER_ACTIONS_XAML,
        width=520,
        min_width=520,
        height=700,
        min_height=700,
        resize_mode=u"NoResize",
        size_to_content_height=False,
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
    def __init__(self, level, checkbox, border):
        self.level = level
        self.checkbox = checkbox
        self.border = border
        self.hover = False


class _CatCheck(object):
    def __init__(self, code, display, checkbox, border, haystack):
        self.code = code
        self.display = display
        self.checkbox = checkbox
        self.border = border
        self.haystack = haystack
        self.visible = True
        self.hover = False


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
        )

        try:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is None:
                mostrar_aviso(host, u"No hay documento activo.")
                return

            doc = uidoc.Document
            result = create_categoria_views(doc, req)
            instruction, content = format_success_dialog(
                result,
                req.categoria_display,
                req.categoria_code,
                req.zona,
                categorias=req.categorias,
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
        self._cat_checks = []
        self._scale_buttons = []
        self._scale_ratio = int(_DEFAULT_VIEW_SCALE)
        self._busy = False

        self._create_handler = _CreateCategoriaViewsHandler()
        self._create_event = ExternalEvent.Create(self._create_handler)

        self._win = XamlReader.Parse(_build_xaml())
        self._panel_categorias = self._win.FindName("PanelCategorias")
        self._txt_cat_count = self._win.FindName("TxtCatCount")
        self._txt_filter = self._win.FindName("TxtFilter")
        self._txt_filter_ph = self._win.FindName("TxtFilterPh")
        self._txt_filter_empty = self._win.FindName("TxtFilterEmpty")
        self._txt_zona = self._win.FindName("TxtZona")
        self._panel_escala = self._win.FindName("PanelEscala")
        self._panel_niveles = self._win.FindName("PanelNiveles")
        self._txt_nivel_count = self._win.FindName("TxtNivelCount")
        self._txt_subtitle = self._win.FindName("TxtSubtitle")
        self._txt_status = self._win.FindName("TxtStatus")
        self._btn_iniciar = self._win.FindName("BtnIniciar")
        self._btn_cancelar = self._win.FindName("BtnCancelar")
        btn_manual = self._win.FindName("BtnManual")
        btn_all = self._win.FindName("BtnSelAll")
        btn_none = self._win.FindName("BtnSelNone")
        btn_cat_all = self._win.FindName("BtnCatSelAll")
        btn_cat_none = self._win.FindName("BtnCatSelNone")

        if self._txt_subtitle is not None:
            try:
                self._txt_subtitle.Text = _SUBTITLE
            except Exception:
                pass

        if self._txt_filter_ph is not None:
            try:
                self._txt_filter_ph.Text = _FILTER_PLACEHOLDER
            except Exception:
                pass
        _sync_filter_placeholder(self._txt_filter, self._txt_filter_ph)

        self._fill_categorias()
        self._apply_categoria_filter()
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
        if self._txt_filter is not None:
            try:
                self._txt_filter.TextChanged += TextChangedEventHandler(
                    self._on_filter_changed
                )
            except Exception:
                pass
        self._refresh_form_state()

        self._btn_iniciar.Click += RoutedEventHandler(self._on_iniciar)
        self._btn_cancelar.Click += RoutedEventHandler(lambda s, e: self._win.Close())
        if btn_manual is not None:
            btn_manual.Click += RoutedEventHandler(
                lambda s, e: _open_manual(self._revit)
            )
        if btn_all is not None:
            btn_all.Click += RoutedEventHandler(
                lambda s, e: self._set_all_levels(True)
            )
        if btn_none is not None:
            btn_none.Click += RoutedEventHandler(
                lambda s, e: self._set_all_levels(False)
            )
        if btn_cat_all is not None:
            btn_cat_all.Click += RoutedEventHandler(
                lambda s, e: self._set_all_categorias(True)
            )
        if btn_cat_none is not None:
            btn_cat_none.Click += RoutedEventHandler(
                lambda s, e: self._set_all_categorias(False)
            )

        self._win.Closed += EventHandler(lambda s, e: singleton.clear())

    def _on_categoria_changed(self, _sender, _e):
        self._refresh_form_state()

    def _on_zona_changed(self, _sender, _e):
        self._refresh_form_state()

    def _on_filter_changed(self, _sender, _e):
        _sync_filter_placeholder(self._txt_filter, self._txt_filter_ph)
        self._apply_categoria_filter()
        self._refresh_form_state()

    def _fill_categorias(self):
        if self._panel_categorias is None:
            return
        self._panel_categorias.Children.Clear()
        self._cat_checks = []
        for code, label in CATEGORIA_OPTIONS:
            name = _categoria_nombre(code, label)
            border, cb, _holder = _make_check_row(
                code, name, False, self._on_categoria_changed, meta_text=code
            )
            haystack = u"{0} {1} {2}".format(code, name, label).lower()
            item = _CatCheck(code, label, cb, border, haystack)
            self._panel_categorias.Children.Add(border)
            self._cat_checks.append(item)

    def _apply_categoria_filter(self):
        q = _filter_query(self._txt_filter)
        visible_n = 0
        for item in self._cat_checks:
            show = (not q) or (q in item.haystack)
            item.visible = show
            if item.border is not None:
                try:
                    item.border.Visibility = (
                        _visibility_visible() if show else _visibility_collapsed()
                    )
                except Exception:
                    pass
            if show:
                visible_n += 1
        if self._txt_filter_empty is not None:
            try:
                total = len(self._cat_checks)
                if total > 0 and visible_n == 0 and q:
                    self._txt_filter_empty.Visibility = _visibility_visible()
                else:
                    self._txt_filter_empty.Visibility = _visibility_collapsed()
            except Exception:
                pass
        return visible_n

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
        if self._panel_niveles is None:
            return
        self._panel_niveles.Children.Clear()
        self._level_checks = []
        levels = _collect_levels_sorted(self._doc)
        for lv in levels:
            try:
                name = _as_unicode(lv.Name or u"")
            except Exception:
                name = u"?"
            border, cb, _holder = _make_check_row(
                lv, name, True, self._on_level_changed
            )
            item = _LevelCheck(lv, cb, border)
            self._panel_niveles.Children.Add(border)
            self._level_checks.append(item)

    def _on_level_changed(self, _sender, _e):
        self._refresh_form_state()

    def _set_all_levels(self, checked):
        for item in self._level_checks:
            item.checkbox.IsChecked = checked
        self._refresh_form_state()

    def _set_all_categorias(self, checked):
        for item in self._cat_checks:
            if not getattr(item, u"visible", True):
                continue
            item.checkbox.IsChecked = checked
        self._refresh_form_state()

    def _get_selected_categorias(self):
        out = []
        for item in self._cat_checks:
            try:
                if item.checkbox.IsChecked:
                    out.append((item.code, item.display))
            except Exception:
                continue
        return out

    def _visible_categoria_count(self):
        n = 0
        for item in self._cat_checks:
            if getattr(item, u"visible", True):
                n += 1
        return n

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

        cats = self._get_selected_categorias()
        zona = self._get_zona()
        n_cat = len(cats)
        total_cat = len(self._cat_checks)
        visible_cat = self._visible_categoria_count()
        q = _filter_query(self._txt_filter)
        if self._txt_cat_count is not None:
            if total_cat <= 0:
                self._txt_cat_count.Text = u"0 de 0 seleccionadas"
            elif q:
                self._txt_cat_count.Text = u"{0} visibles · {1} seleccionadas".format(
                    visible_cat, n_cat
                )
            else:
                self._txt_cat_count.Text = u"{0} de {1} seleccionadas".format(
                    n_cat, total_cat
                )

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

        from vistas_por_categoria.service import plans_per_level_for_categoria

        scale = self._get_selected_scale()
        can_run = n_cat > 0 and n > 0
        _style_crear_button(self._btn_iniciar, can_run)

        plants = sum(plans_per_level_for_categoria(c[0]) for c in cats) * n
        if n_cat <= 0:
            self._set_status(u"Marque una o más categorías.")
        elif n == 0:
            self._set_status(u"Seleccione al menos un nivel.")
        elif n_cat == 1:
            self._set_status(
                u"{0} · {1} · 1:{2} · {3} plantas".format(
                    cats[0][0], zona, scale, plants
                )
            )
        else:
            self._set_status(
                u"{0} cat. · {1} · 1:{2} · {3} plantas".format(
                    n_cat, zona, scale, plants
                )
            )

    def _on_iniciar(self, sender, args):
        if self._busy:
            return

        cats = self._get_selected_categorias()
        if not cats:
            self._set_status(u"Marque una o más categorías.")
            return

        levels = self._get_selected_levels()
        if not levels:
            self._set_status(u"Seleccione al menos un nivel.")
            return

        from vistas_por_categoria.service import VistasPorCategoriaRequest

        scale = self._get_selected_scale()
        zona = self._get_zona()
        req = VistasPorCategoriaRequest(cats, zona, scale, list(levels))

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
