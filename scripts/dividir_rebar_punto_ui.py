# -*- coding: utf-8 -*-
"""UI WPF — Dividir y Traslapar (multipunto, traslape por diámetro)."""

from __future__ import print_function

import os
import weakref

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System import Action, AppDomain, Byte, Double, EventHandler
from System.Collections.Generic import List as DotNetList
from System.Windows import (
    FontWeights,
    HorizontalAlignment,
    Point,
    RoutedEventHandler,
    Size,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    WindowState,
)
from System.Windows.Controls import (
    Canvas,
    ComboBoxItem,
    Orientation,
    SelectionChangedEventHandler,
    StackPanel,
    TextBlock,
    TextBox,
)
from System.Windows.Input import (
    Cursors,
    Key,
    KeyEventHandler,
    Mouse,
    MouseButton,
    MouseButtonEventHandler,
    MouseButtonState,
    MouseEventHandler,
    MouseWheelEventHandler,
)
from System.Windows.Markup import XamlReader
from System.Windows.Media import (
    Color,
    EdgeMode,
    FontFamily,
    PenLineCap,
    PenLineJoin,
    PointCollection,
    RenderOptions,
    SolidColorBrush,
)
from System.Windows.Shapes import Line, Polyline, Rectangle
from System.Windows.Threading import DispatcherPriority
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

from bimtools_ui_tokens import (
    ACCENT_PRIMARY,
    BTN_MANUAL,
    FG_BODY,
    FG_MUTED,
    FONT_SIZE_HINT,
)
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from dividir_rebar_punto import _TITULO, _as_unicode, grade_from_combo_text, mostrar_aviso
from dividir_rebar_punto_core import (
    DividirRebarProgress,
    divide_rebar_at_cuts,
    lap_mm_for_rebar,
    layout_label,
    pick_rebar,
    prepare_division_session,
    resolve_active_model_view,
)
from dividir_rebar_punto_geom import (
    build_spans_mm,
    ceil_mm_to_nearest_10,
    compute_canvas_mapping,
    fabricated_lengths_mm,
    fit_polyline_to_canvas,
    lap_zone_around_cut,
    map_polyline_uv_to_canvas_px,
    map_uv_to_canvas_px,
    nearest_arc_length_px,
    normalize_lap_mode,
    piece_intervals_with_lap,
    point_at_arc_length_uv,
    set_span_fabricated_length_mm,
    span_index_at_mm,
    span_midpoint_mm,
    tangent_at_arc_length_uv,
    tramo_label,
    validate_cuts_with_lap,
    LAP_MODE_ENDPOINT_NEXT,
    LAP_MODE_ENDPOINT_PREV,
    LAP_MODE_SYMMETRIC,
)
from revit_wpf_window_position import revit_main_hwnd

_SINGLETON_KEY = u"Arainco_DividirRebarPuntoTraslape_UI"
_MIN_PIECE_MM = 100.0
# Retiene el controlador mientras corre el ExternalEvent tras cerrar la UI.
_APPLY_HOLD = None

_CANVAS_MIN_W = 420.0
_CANVAS_MIN_H = 480.0
# Alias de mínimos (fallback antes del primer layout).
_CANVAS_W = _CANVAS_MIN_W
_CANVAS_H = _CANVAS_MIN_H
_MARGIN = 40.0
# Rail derecho (controles) — misma idea que SECTION_RAIL en Armado Vigas.
_RAIL_W = 420
# Ventana base (antes de maximizar); misma familia que Armado Vigas.
_WINDOW_W = 1360
_WINDOW_H = 920
_WINDOW_MIN_W = 960
_WINDOW_MIN_H = 640
_ZOOM_MIN = 0.5
_ZOOM_MAX = 8.0
_ZOOM_STEP = 1.15
_ZOOM_DEFAULT = 1.0

_BAR_THICK = 2.0
_CUT_HIT_PX = 10.0
_CUT_MARK_PX = 14.0
_CUT_ADD_HIT_PX = 18.0

_COLOR_BAR = ACCENT_PRIMARY
# Hormigón en alzado/sección — misma lectura que Armado Vigas (gris, no azul barra).
_COLOR_CONCRETE_EDGE = u"#9ca3a8"
_COLOR_CONCRETE_FILL_RGB = (138, 142, 148)
_COLOR_CONCRETE_FILL_A = 148
_COLOR_CONCRETE_EDGE_A = 175
_COLOR_CONCRETE_STROKE = 0.55
_COLOR_CUT = u"#f87171"
_COLOR_LAP = u"#fbbf24"
# Colores compartidos editor T# ↔ tramo en canvas (cíclico).
_TRAMO_PALETTE = (
    u"#38bdf8",
    u"#a78bfa",
    u"#34d399",
    u"#fb923c",
    u"#f472b6",
    u"#facc15",
)
_TRAMO_HIT_PX = 14.0
# Muestreo denso → curvas menos angulosas en el esquema.
_SPAN_SAMPLE_MM = 28.0

_BRUSH_CACHE = {}
_XAML_CACHE = None
_XAML_AD_KEY = u"Arainco_DividirRebarXamlCache_v23_canvas_zoom"
# Claves viejas: forzar XAML fresco tras cambios de panel / shell.
_XAML_AD_KEYS_PURGE = (
    u"Arainco_DividirRebarXamlCache_v4_tight_height",
    u"Arainco_DividirRebarXamlCache_v5_no_cuts_edit",
    u"Arainco_DividirRebarXamlCache_v6_tramos_edit",
    u"Arainco_DividirRebarXamlCache_v7_tramos_host",
    u"Arainco_DividirRebarXamlCache_v8_pick_preview",
    u"Arainco_DividirRebarXamlCache_v9_pickobjects_finish",
    u"Arainco_DividirRebarXamlCache_v10_tramo_link",
    u"Arainco_DividirRebarXamlCache_v15_wpf_shell",
    u"Arainco_DividirRebarXamlCache_v16_footer_static",
    u"Arainco_DividirRebarXamlCache_v18_manual_ligero",
    u"Arainco_DividirRebarXamlCache_v19_shell_actual",
    u"Arainco_DividirRebarXamlCache_v20_canvas_cuts",
    u"Arainco_DividirRebarXamlCache_v21_armado_shell",
    u"Arainco_DividirRebarXamlCache_v22_fill_canvas",
    _XAML_AD_KEY,
)


def _tramo_color(span_index):
    if not _TRAMO_PALETTE:
        return _COLOR_BAR
    try:
        i = int(span_index)
    except Exception:
        i = 0
    return _TRAMO_PALETTE[i % len(_TRAMO_PALETTE)]


def _purge_xaml_cache():
    global _XAML_CACHE
    _XAML_CACHE = None
    for key in _XAML_AD_KEYS_PURGE:
        try:
            AppDomain.CurrentDomain.SetData(key, None)
        except Exception:
            pass


_purge_xaml_cache()


def _brush(hex_color, alpha=255):
    key = (u"{0}|{1}".format(_as_unicode(hex_color), int(alpha)))
    cached = _BRUSH_CACHE.get(key)
    if cached is not None:
        return cached
    h = _as_unicode(hex_color).lstrip(u"#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    br = SolidColorBrush(Color.FromArgb(Byte(alpha), Byte(r), Byte(g), Byte(b)))
    try:
        br.Freeze()
    except Exception:
        pass
    _BRUSH_CACHE[key] = br
    return br


def _brush_rgb(r, g, b, alpha=255):
    key = (u"rgb|{0}|{1}|{2}|{3}".format(int(r), int(g), int(b), int(alpha)))
    cached = _BRUSH_CACHE.get(key)
    if cached is not None:
        return cached
    br = SolidColorBrush(
        Color.FromArgb(Byte(alpha), Byte(r), Byte(g), Byte(b))
    )
    try:
        br.Freeze()
    except Exception:
        pass
    _BRUSH_CACHE[key] = br
    return br


# Shell alineado con Armado Vigas: cinta blanca + header + hint + canvas|rail + footer.
_XAML_DIVIDIR = u"""<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  xmlns:po="http://schemas.microsoft.com/winfx/2006/xaml/presentation/options"
  Title="Arainco"
  Height="__WINDOW_H__" Width="__WINDOW_W__"
  MinHeight="__WINDOW_MIN_H__" MinWidth="__WINDOW_MIN_W__"
  ResizeMode="CanResize"
  WindowStartupLocation="Manual"
  Background="#071018"
  FontFamily="Segoe UI"
  FontSize="12"
  ShowInTaskbar="False">
  <Window.Resources>
__BIMTOOLS_DARK_STYLES__
    <SolidColorBrush x:Key="DividirAppBg" Color="#071018" po:Freeze="True"/>
    <SolidColorBrush x:Key="DividirPanelBg" Color="#0a1620" po:Freeze="True"/>
    <SolidColorBrush x:Key="DividirInputBg" Color="#050E18" po:Freeze="True"/>
    <SolidColorBrush x:Key="DividirBorder" Color="#21465C" po:Freeze="True"/>
    <SolidColorBrush x:Key="DividirFgHi" Color="#E8F4F8" po:Freeze="True"/>
    <SolidColorBrush x:Key="DividirFgMid" Color="#95B8CC" po:Freeze="True"/>
    <SolidColorBrush x:Key="DividirFgLo" Color="#64748b" po:Freeze="True"/>
    <SolidColorBrush x:Key="DividirAccent" Color="#5BC0DE" po:Freeze="True"/>
  </Window.Resources>
  <Border Background="{StaticResource DividirAppBg}" BorderBrush="{StaticResource DividirBorder}"
          BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Grid Grid.Row="0" Margin="0,0,0,8">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="TxtTitle" Grid.Row="0" Text="Arainco: Dividir y Traslapar"
                   Foreground="{StaticResource DividirFgHi}" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Grid.Row="1" Margin="0,6,0,0"
                   Foreground="{StaticResource DividirFgMid}" TextWrapping="Wrap"
                   Text="Cortes en el esquema · traslape por diámetro · largos fabricados"/>
      </Grid>

      <TextBlock x:Name="TxtHint" Grid.Row="1"
                 Foreground="{StaticResource DividirFgLo}" FontSize="10"
                 TextWrapping="Wrap" Margin="0,0,0,10"
                 Text="Clic en la barra añade un corte · rueda acerca/aleja · botón medio desplaza · clic en C# lo quita · clic en T# enfoca el editor."/>

      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="__RAIL_W__"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="{StaticResource DividirPanelBg}"
                BorderBrush="{StaticResource DividirBorder}" BorderThickness="1"
                CornerRadius="4,0,0,4" Padding="0">
          <Grid Margin="10,8,10,8">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0" Margin="0,0,0,6">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="Esquema de alzado"
                         Foreground="{StaticResource DividirFgMid}"
                         FontSize="11" FontWeight="SemiBold"
                         VerticalAlignment="Center"/>
              <StackPanel Grid.Column="1" Orientation="Horizontal"
                          VerticalAlignment="Center">
                <Button x:Name="BtnZoomOut" Content="−"
                        Style="{StaticResource BtnSelectOutline}"
                        MinWidth="28" Height="22" Padding="6,1"
                        ToolTip="Alejar (rueda del ratón)"/>
                <TextBlock x:Name="TxtZoom" Text="100%"
                           Foreground="{StaticResource DividirAccent}"
                           FontSize="11" FontWeight="Bold" MinWidth="40"
                           TextAlignment="Center" VerticalAlignment="Center"
                           Margin="6,0,6,0"/>
                <Button x:Name="BtnZoomIn" Content="+"
                        Style="{StaticResource BtnSelectOutline}"
                        MinWidth="28" Height="22" Padding="6,1"
                        ToolTip="Acercar (rueda del ratón)"/>
                <Button x:Name="BtnZoomFit" Content="100%"
                        Style="{StaticResource BtnSelectOutline}"
                        MinWidth="48" Height="22" Padding="8,1" Margin="6,0,0,0"
                        ToolTip="Ajustar el esquema a la vista"/>
              </StackPanel>
            </Grid>
            <Border x:Name="BdrCanvasHost" Grid.Row="1"
                    Background="{StaticResource DividirInputBg}"
                    BorderThickness="0" ClipToBounds="True"
                    HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
              <Canvas x:Name="CnvBar"
                      Background="{StaticResource DividirInputBg}" Cursor="Arrow"
                      ClipToBounds="False" Focusable="True"
                      HorizontalAlignment="Left" VerticalAlignment="Top"
                      SnapsToDevicePixels="True"
                      RenderOptions.EdgeMode="Aliased"/>
            </Border>
          </Grid>
        </Border>

        <Border Grid.Column="1" Background="{StaticResource DividirPanelBg}"
                BorderBrush="{StaticResource DividirBorder}" BorderThickness="1,1,1,1"
                CornerRadius="0,4,4,0" Padding="8,8">
          <ScrollViewer VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Disabled"
                        CanContentScroll="False"
                        IsDeferredScrollingEnabled="True">
            <StackPanel>
              <TextBlock Style="{StaticResource LabelSmall}"
                         Text="Grado de hormigón (tabla traslape)" Margin="0,0,0,4"/>
              <Grid Margin="0,0,0,10">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <ComboBox x:Name="CmbGrade" Grid.Column="0"
                          Style="{StaticResource Combo}" Height="28"/>
                <Button x:Name="BtnLimpiar" Grid.Column="1" Content="Limpiar cortes"
                        Style="{StaticResource BtnSelectOutline}" MinWidth="110"
                        Margin="8,0,0,0"/>
              </Grid>

              <Border Background="{StaticResource DividirAppBg}"
                      BorderBrush="{StaticResource DividirBorder}" BorderThickness="1"
                      CornerRadius="4" Padding="10" Margin="0,0,0,10">
                <TextBlock x:Name="TxtMeta" TextWrapping="Wrap"
                           Foreground="{StaticResource DividirFgMid}"
                           FontSize="11" Text=""/>
              </Border>

              <TextBlock x:Name="TxtAlert" TextWrapping="Wrap" FontSize="11"
                         Foreground="{StaticResource DividirFgMid}" Margin="0,0,0,10"/>

              <TextBlock Text="Modo de traslape" Style="{StaticResource LabelSmall}"
                         Margin="0,0,0,4"/>
              <ComboBox x:Name="CmbLapMode" Style="{StaticResource Combo}"
                        Height="28" Margin="0,0,0,4"/>
              <TextBlock x:Name="TxtLapModeHint" TextWrapping="Wrap"
                         Foreground="{StaticResource DividirFgLo}" FontSize="10"
                         Margin="0,0,0,10" Text=""/>

              <TextBlock Text="Largos fabricados (mm)" Style="{StaticResource LabelSmall}"
                         Margin="0,0,0,4"/>
              <Border x:Name="BrdTramosHost" Background="Transparent"
                      MinHeight="96" Margin="0,0,0,4"/>
            </StackPanel>
          </ScrollViewer>
        </Border>
      </Grid>

      <Grid Grid.Row="3" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="BtnManual" Grid.Column="0" Content="Manual"
                Style="{StaticResource BtnSelectOutline}"
                Background="__BTN_MANUAL__" MinWidth="96" Padding="8,2"
                ToolTip="Abrir manual de usuario" VerticalAlignment="Center"
                Margin="0,0,12,0"/>
        <TextBlock x:Name="TxtStatus" Grid.Column="1" VerticalAlignment="Center"
                   Foreground="{StaticResource DividirFgLo}" FontSize="10"
                   TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnClose" Content="Cerrar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnApply" Content="Aplicar"
                  Style="{StaticResource BtnPrimary}" MinWidth="160"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>"""


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
        # scripts/ → BIMTools.extension/
        ext_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        for root, dirs, files in os.walk(ext_dir):
            base = os.path.basename(root)
            if base.endswith(u"DividirRebarPuntoTraslape.pushbutton"):
                candidates.append(os.path.join(root, u"manual_usuario.html"))
                dirs[:] = []
                continue
            # No entrar en carpetas pesadas / irrelevantes.
            skip = []
            for d in dirs:
                if d in (u".git", u"__pycache__", u"node_modules", u"canvases"):
                    skip.append(d)
            for d in skip:
                try:
                    dirs.remove(d)
                except Exception:
                    pass
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
            u"Debe estar en la carpeta del pushbutton de la herramienta.",
        )
        return
    try:
        os.startfile(path)
    except Exception as ex:
        mostrar_aviso(
            uiapp,
            u"No se pudo abrir el manual.",
            _as_unicode(ex),
        )


def _build_xaml():
    global _XAML_CACHE
    if _XAML_CACHE:
        return _XAML_CACHE
    try:
        cached = AppDomain.CurrentDomain.GetData(_XAML_AD_KEY)
        if cached:
            _XAML_CACHE = cached
            return _XAML_CACHE
    except Exception:
        pass
    xaml = _XAML_DIVIDIR.replace(u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)
    xaml = xaml.replace(u"__WINDOW_W__", u"{0}".format(int(_WINDOW_W)))
    xaml = xaml.replace(u"__WINDOW_H__", u"{0}".format(int(_WINDOW_H)))
    xaml = xaml.replace(u"__WINDOW_MIN_W__", u"{0}".format(int(_WINDOW_MIN_W)))
    xaml = xaml.replace(u"__WINDOW_MIN_H__", u"{0}".format(int(_WINDOW_MIN_H)))
    xaml = xaml.replace(u"__RAIL_W__", u"{0:.0f}".format(float(_RAIL_W)))
    xaml = xaml.replace(u"__BTN_MANUAL__", _as_unicode(BTN_MANUAL))
    _XAML_CACHE = xaml
    try:
        AppDomain.CurrentDomain.SetData(_XAML_AD_KEY, _XAML_CACHE)
    except Exception:
        pass
    return _XAML_CACHE


def _prepare_window(win, uiapp):
    """Misma preparación de monitor/maximizado que Armado Vigas."""
    if win is None:
        return
    try:
        from System import Double
        from System.Windows import SizeToContent, WindowStartupLocation
        from revit_wpf_window_position import (
            bind_maximize_wpf_on_revit_monitor,
            preposition_wpf_window_on_work_area,
            _monitor_work_area_px,
            _primary_work_area_px,
        )

        hwnd = revit_main_hwnd(uiapp)
        try:
            win.MaxWidth = Double.PositiveInfinity
            win.MaxHeight = Double.PositiveInfinity
        except Exception:
            pass
        try:
            win.SizeToContent = SizeToContent.Manual
        except Exception:
            pass
        try:
            win.WindowStartupLocation = WindowStartupLocation.Manual
        except Exception:
            pass
        try:
            area = _monitor_work_area_px(hwnd)
            if area is None:
                area = _primary_work_area_px()
            if area is not None:
                preposition_wpf_window_on_work_area(
                    win, area[0], area[1], area[2], area[3], hwnd,
                )
        except Exception:
            pass
        bind_maximize_wpf_on_revit_monitor(win, hwnd)
    except Exception:
        pass


def _maximize_window(win):
    """Despliega la herramienta maximizada."""
    if win is None:
        return
    try:
        from System.Windows import SizeToContent

        win.SizeToContent = SizeToContent.Manual
    except Exception:
        pass
    try:
        win.WindowState = WindowState.Maximized
    except Exception:
        pass


def _native_bring_to_front(win):
    """
    Foco nativo ligero. Evitar AttachThreadInput + keybd_event(ALT):
    con el message pump de Revit suelen provocar micro-cuelgues / z-order raro.
    """
    if win is None:
        return
    try:
        from System.Windows.Interop import WindowInteropHelper
        import ctypes

        helper = WindowInteropHelper(win)
        hwnd = helper.Handle
        if hwnd is None or int(hwnd) == 0:
            try:
                hwnd = helper.EnsureHandle()
            except Exception:
                return
        if hwnd is None or int(hwnd) == 0:
            return
        hwnd = int(hwnd)
        user32 = ctypes.windll.user32
        try:
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
        except Exception:
            pass
    except Exception:
        pass


def _bring_window_to_front(win):
    if win is None:
        return
    try:
        win.ShowActivated = True
    except Exception:
        pass
    try:
        _maximize_window(win)
    except Exception:
        pass
    try:
        win.Topmost = True
        win.Activate()
        win.Focus()
    except Exception:
        pass
    _native_bring_to_front(win)
    try:
        win.Topmost = False
        win.Activate()
    except Exception:
        pass


def _existing_window():
    try:
        return AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        return None


def _set_singleton(win):
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass


def _clear_singleton():
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
    except Exception:
        pass


def _focus_existing(uiapp):
    holder = _existing_window()
    if holder is None:
        return False
    win = getattr(holder, u"_win", None) or holder
    try:
        if win is None or not win.IsLoaded:
            _clear_singleton()
            return False
        _maximize_window(win)
        try:
            if float(getattr(win, u"Opacity", 1.0) or 1.0) < 0.5:
                win.Opacity = 1.0
        except Exception:
            pass
        if not win.IsVisible:
            win.Show()
        _bring_window_to_front(win)
        mostrar_aviso(uiapp, u"La herramienta ya está en ejecución.")
        return True
    except Exception:
        _clear_singleton()
        return False


def _clear_selection(uidoc):
    """Quita la selección activa para no activar la cinta Modify de Revit."""
    if uidoc is None:
        return
    try:
        uidoc.Selection.SetElementIds(DotNetList[ElementId]())
    except Exception:
        pass


class _ApplyHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def GetName(self):
        return u"DividirRebarPunto.Apply"

    def Execute(self, uiapp):
        global _APPLY_HOLD
        win = None
        try:
            win = self._window_ref() if self._window_ref is not None else None
        except Exception:
            win = None
        if win is None:
            win = _APPLY_HOLD
        try:
            if win is not None:
                win.apply_in_revit(uiapp)
        except Exception as ex:
            try:
                mostrar_aviso(uiapp, u"No se pudo dividir.", _as_unicode(ex))
            except Exception:
                pass
        finally:
            _APPLY_HOLD = None
            try:
                if win is not None:
                    win._finish_after_apply()
            except Exception:
                _clear_singleton()


class DividirRebarPuntoWindow(object):
    def __init__(self, uiapp, rebar, session):
        self._uiapp = uiapp
        xaml = _build_xaml()
        self._win = XamlReader.Parse(xaml)
        self._session = session or {}
        self._rebar_id = self._session.get(u"rebar_id")
        self._total_mm = float(self._session.get(u"total_mm") or 0.0)
        self._lap_mm = float(self._session.get(u"lap_mm") or 0.0)
        self._diameter_mm = self._session.get(u"diameter_mm")
        self._cuts_mm = []
        self._lap_mode = LAP_MODE_SYMMETRIC
        self._active_tramo_index = -1
        self._tramo_editors = {}
        self._scale = 1.0
        self._view_zoom = float(_ZOOM_DEFAULT)
        self._view_pan_x = 0.0
        self._view_pan_y = 0.0
        self._panning = False
        self._pan_last = None
        self._pan_origin_cursor = None
        self._canvas_w = float(_CANVAS_MIN_W)
        self._canvas_h = float(_CANVAS_MIN_H)
        self._canvas_host = None
        self._resize_token = 0
        self._busy = False
        self._rebuilding = False
        self._schema_built = False
        self._closing_for_apply = False
        self._apply_started = False
        self._plan_uv = list(self._session.get(u"plan_points_uv") or [])
        self._plan_arc_mm = list(self._session.get(u"plan_arc_mm") or [])
        self._plan_flip_v = bool(self._session.get(u"plan_flip_v", True))
        self._plan_px = []
        self._context_polylines_uv = list(
            self._session.get(u"context_polylines_uv") or []
        )
        self._context_fill_rects_uv = list(
            self._session.get(u"context_fill_rects_uv") or []
        )
        self._context_n_elems = int(self._session.get(u"context_n_elems") or 0)
        self._context_n_polylines = int(self._session.get(u"context_n_polylines") or 0)
        self._canvas_mapping = None
        if len(self._plan_uv) < 2:
            L = max(1.0, self._total_mm)
            self._plan_uv = [[0.0, 0.0], [L, 0.0]]
            self._plan_arc_mm = [0.0, L]
            self._plan_flip_v = False
        elif not self._plan_arc_mm or len(self._plan_arc_mm) != len(self._plan_uv):
            from dividir_rebar_punto_geom import polyline_arc_lengths

            arcs = polyline_arc_lengths(self._plan_uv)
            tot = float(arcs[-1]) if arcs else 0.0
            if tot > 1e-6 and self._total_mm > 1e-6:
                ratio = self._total_mm / tot
                self._plan_arc_mm = [float(a) * ratio for a in arcs]
            else:
                self._plan_arc_mm = arcs

        self._txt_title = self._win.FindName(u"TxtTitle")
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_meta = self._win.FindName(u"TxtMeta")
        self._txt_alert = self._win.FindName(u"TxtAlert")
        self._txt_status = self._win.FindName(u"TxtStatus")
        self._cmb_grade = self._win.FindName(u"CmbGrade")
        self._cmb_lap_mode = self._win.FindName(u"CmbLapMode")
        self._txt_lap_mode_hint = self._win.FindName(u"TxtLapModeHint")
        self._cnv = self._win.FindName(u"CnvBar")
        self._canvas_host = self._win.FindName(u"BdrCanvasHost")
        self._btn_zoom_in = self._win.FindName(u"BtnZoomIn")
        self._btn_zoom_out = self._win.FindName(u"BtnZoomOut")
        self._btn_zoom_fit = self._win.FindName(u"BtnZoomFit")
        self._txt_zoom = self._win.FindName(u"TxtZoom")
        if self._cnv is not None:
            try:
                self._cnv.SnapsToDevicePixels = False
                self._cnv.UseLayoutRounding = False
            except Exception:
                pass
            try:
                RenderOptions.SetEdgeMode(self._cnv, EdgeMode.Unspecified)
            except Exception:
                pass
            try:
                self._cnv.Width = float(self._canvas_w)
                self._cnv.Height = float(self._canvas_h)
            except Exception:
                pass
        self._brd_tramos_host = self._win.FindName(u"BrdTramosHost")
        self._pnl_tramos = None
        self._ensure_tramos_panel()
        self._btn_limpiar = self._win.FindName(u"BtnLimpiar")
        self._btn_apply = self._win.FindName(u"BtnApply")
        self._btn_close = self._win.FindName(u"BtnClose")
        self._btn_manual = self._win.FindName(u"BtnManual")

        if self._txt_title is not None:
            self._txt_title.Text = _TITULO
        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = (
                u"Clic en el esquema de alzado para marcar cortes; rueda para "
                u"zoom y botón medio para desplazar. Afine el largo fabricado "
                u"de cada tramo y pulse Aplicar."
            )

        self._handler_apply = _ApplyHandler(weakref.ref(self))
        self._ext_apply = ExternalEvent.Create(self._handler_apply)

        self._populate_grade_combo()
        self._populate_lap_mode_combo()
        self._wire()
        self._update_meta()
        self._update_lap_mode_hint()

        # Esquema / paneles se montan en show() antes de Show (sin shell vacío).
        self._ui_set_alert(
            u"Clic en el esquema de alzado (cerca de la barra) para añadir cortes. "
            u"Clic cerca de C# para quitarlos. Rueda: zoom · botón medio: desplazar.",
            u"info",
        )

    def _populate_grade_combo(self):
        if self._cmb_grade is None:
            return
        self._cmb_grade.Items.Clear()
        items = [
            u"G25",
            u"G35",
            u"G45",
        ]
        for i, label in enumerate(items):
            item = ComboBoxItem()
            item.Content = label
            self._cmb_grade.Items.Add(item)
        self._cmb_grade.SelectedIndex = 0  # G25 por defecto

    def _populate_lap_mode_combo(self):
        if self._cmb_lap_mode is None:
            return
        self._cmb_lap_mode.Items.Clear()
        options = (
            (LAP_MODE_SYMMETRIC, u"Simétrico ± L/2 (actual)"),
            (LAP_MODE_ENDPOINT_PREV, u"Endpoint · estira tramo anterior (+L)"),
            (LAP_MODE_ENDPOINT_NEXT, u"Endpoint · estira tramo siguiente (−L)"),
        )
        selected = 0
        for i, (key, label) in enumerate(options):
            item = ComboBoxItem()
            item.Content = label
            item.Tag = key
            self._cmb_lap_mode.Items.Add(item)
            if key == self._lap_mode:
                selected = i
        self._cmb_lap_mode.SelectedIndex = selected

    def _selected_lap_mode(self):
        try:
            item = self._cmb_lap_mode.SelectedItem if self._cmb_lap_mode else None
            if item is not None and getattr(item, u"Tag", None):
                return normalize_lap_mode(item.Tag)
        except Exception:
            pass
        return normalize_lap_mode(self._lap_mode)

    def _lap_mode_hint_text(self, mode=None):
        m = normalize_lap_mode(mode if mode is not None else self._selected_lap_mode())
        if m == LAP_MODE_ENDPOINT_PREV:
            return (
                u"Todo el solape L sale del extremo del tramo que llega al corte "
                u"(termina en C+L)."
            )
        if m == LAP_MODE_ENDPOINT_NEXT:
            return (
                u"Todo el solape L sale del extremo del tramo que parte del corte "
                u"(empieza en C−L)."
            )
        return u"Cada barra aporta L/2 al solape, centrado en el corte."

    def _update_lap_mode_hint(self):
        if self._txt_lap_mode_hint is None:
            return
        self._txt_lap_mode_hint.Text = self._lap_mode_hint_text()

    def _wire(self):
        self._btn_close.Click += RoutedEventHandler(self._on_close)
        self._btn_apply.Click += RoutedEventHandler(self._on_apply)
        self._btn_limpiar.Click += RoutedEventHandler(self._on_limpiar)
        if self._btn_manual is not None:
            self._btn_manual.Click += RoutedEventHandler(self._on_manual)
        self._cnv.MouseLeftButtonDown += MouseButtonEventHandler(self._on_canvas_click)
        self._wire_canvas_nav()
        self._wire_zoom_chrome()
        self._sync_zoom_label()
        self._win.Closed += EventHandler(self._on_closed)
        try:
            from System.Windows import SizeChangedEventHandler as _SCEH

            self._win.SizeChanged += _SCEH(self._on_window_size_changed)
        except Exception:
            try:
                self._win.SizeChanged += RoutedEventHandler(self._on_window_size_changed)
            except Exception:
                pass
        try:
            if self._canvas_host is not None:
                from System.Windows import SizeChangedEventHandler as _SCEH2

                self._canvas_host.SizeChanged += _SCEH2(self._on_window_size_changed)
        except Exception:
            pass
        try:
            self._cmb_grade.SelectionChanged += SelectionChangedEventHandler(
                self._on_grade_changed
            )
        except Exception:
            pass
        try:
            if self._cmb_lap_mode is not None:
                self._cmb_lap_mode.SelectionChanged += SelectionChangedEventHandler(
                    self._on_lap_mode_changed
                )
        except Exception:
            pass

    def _wire_zoom_chrome(self):
        if self._btn_zoom_in is not None:
            try:
                self._btn_zoom_in.Click += RoutedEventHandler(
                    lambda s, e: self._zoom_view_by(_ZOOM_STEP)
                )
            except Exception:
                pass
        if self._btn_zoom_out is not None:
            try:
                self._btn_zoom_out.Click += RoutedEventHandler(
                    lambda s, e: self._zoom_view_by(1.0 / _ZOOM_STEP)
                )
            except Exception:
                pass
        if self._btn_zoom_fit is not None:
            try:
                self._btn_zoom_fit.Click += RoutedEventHandler(
                    lambda s, e: self._reset_view_zoom()
                )
            except Exception:
                pass

    def _wire_canvas_nav(self):
        host = self._canvas_host
        cnv = self._cnv
        if host is not None:
            try:
                host.PreviewMouseWheel += MouseWheelEventHandler(self._on_canvas_wheel)
            except Exception:
                try:
                    host.MouseWheel += MouseWheelEventHandler(self._on_canvas_wheel)
                except Exception:
                    pass
            try:
                host.PreviewMouseDown += MouseButtonEventHandler(self._on_canvas_pan_down)
                host.PreviewMouseMove += MouseEventHandler(self._on_canvas_pan_move)
                host.PreviewMouseUp += MouseButtonEventHandler(self._on_canvas_pan_up)
            except Exception:
                pass
            try:
                host.LostMouseCapture += MouseEventHandler(self._on_canvas_pan_lost)
            except Exception:
                pass
        if cnv is not None:
            try:
                cnv.Focusable = True
            except Exception:
                pass
            try:
                cnv.MouseWheel += MouseWheelEventHandler(self._on_canvas_wheel)
            except Exception:
                pass

    def _clamp_view_zoom(self, z):
        try:
            z = float(z)
        except Exception:
            z = float(_ZOOM_DEFAULT)
        if z < float(_ZOOM_MIN):
            return float(_ZOOM_MIN)
        if z > float(_ZOOM_MAX):
            return float(_ZOOM_MAX)
        return z

    def _view_zoom_label_text(self):
        try:
            return u"{0:.0f}%".format(self._clamp_view_zoom(self._view_zoom) * 100.0)
        except Exception:
            return u"100%"

    def _sync_zoom_label(self):
        try:
            if self._txt_zoom is not None:
                self._txt_zoom.Text = self._view_zoom_label_text()
        except Exception:
            pass

    def _canvas_center_px(self):
        return float(self._canvas_w) * 0.5, float(self._canvas_h) * 0.5

    def _reset_view_zoom(self, announce=True):
        self._view_zoom = float(_ZOOM_DEFAULT)
        self._view_pan_x = 0.0
        self._view_pan_y = 0.0
        self._sync_zoom_label()
        try:
            self._redraw_canvas(detail=True)
        except Exception:
            pass
        if announce:
            self._ui_set_status(u"Zoom alzado · 100 % (ajuste a la vista).")

    def _set_view_zoom_at(self, zoom, pivot_x, pivot_y, announce=True):
        z_old = self._clamp_view_zoom(self._view_zoom)
        z_new = self._clamp_view_zoom(zoom)
        if abs(z_new - z_old) < 1e-9:
            self._sync_zoom_label()
            return
        ratio = z_new / z_old if z_old > 1e-9 else 1.0
        try:
            mx = float(pivot_x)
            my = float(pivot_y)
        except Exception:
            mx, my = self._canvas_center_px()
        cx, cy = self._canvas_center_px()
        self._view_pan_x = (mx - cx) * (1.0 - ratio) + float(self._view_pan_x) * ratio
        self._view_pan_y = (my - cy) * (1.0 - ratio) + float(self._view_pan_y) * ratio
        self._view_zoom = z_new
        self._sync_zoom_label()
        try:
            self._redraw_canvas(detail=True)
        except Exception:
            pass
        if announce:
            self._ui_set_status(
                u"Zoom alzado · {0}".format(self._view_zoom_label_text())
            )

    def _zoom_view_by(self, factor, pivot=None):
        try:
            step = float(factor)
        except Exception:
            return
        if step <= 0:
            return
        if pivot is None:
            px, py = self._canvas_center_px()
        else:
            try:
                px, py = float(pivot[0]), float(pivot[1])
            except Exception:
                px, py = self._canvas_center_px()
        self._set_view_zoom_at(float(self._view_zoom) * step, px, py)

    def _on_canvas_wheel(self, sender, args):
        if self._busy:
            return
        try:
            if args.Handled:
                return
        except Exception:
            pass
        try:
            delta = int(args.Delta)
        except Exception:
            return
        if delta == 0:
            return
        pos_el = self._cnv if self._cnv is not None else self._canvas_host
        try:
            pos = args.GetPosition(pos_el)
            mx = float(pos.X)
            my = float(pos.Y)
        except Exception:
            mx, my = self._canvas_center_px()
        factor = float(_ZOOM_STEP) if delta > 0 else (1.0 / float(_ZOOM_STEP))
        self._zoom_view_by(factor, pivot=(mx, my))
        try:
            args.Handled = True
        except Exception:
            pass

    def _is_middle_button(self, args):
        try:
            return args.ChangedButton == MouseButton.Middle
        except Exception:
            pass
        try:
            return bool(args.MiddleButton) and str(args.MiddleButton).endswith(
                u"Pressed"
            )
        except Exception:
            return False

    def _nav_capture_target(self):
        if self._canvas_host is not None:
            return self._canvas_host
        return self._cnv

    def _on_canvas_pan_down(self, sender, args):
        if not self._is_middle_button(args):
            return
        target = self._nav_capture_target()
        if target is None:
            return
        try:
            self._panning = True
            self._pan_last = args.GetPosition(target)
            try:
                self._pan_origin_cursor = target.Cursor
            except Exception:
                self._pan_origin_cursor = None
            try:
                target.Cursor = Cursors.SizeAll
            except Exception:
                try:
                    target.Cursor = Cursors.Hand
                except Exception:
                    pass
            try:
                target.CaptureMouse()
            except Exception:
                pass
            try:
                args.Handled = True
            except Exception:
                pass
        except Exception:
            self._panning = False
            self._pan_last = None

    def _on_canvas_pan_move(self, sender, args):
        if not self._panning or self._pan_last is None:
            return
        target = self._nav_capture_target()
        if target is None:
            return
        try:
            try:
                if Mouse.MiddleButton != MouseButtonState.Pressed:
                    self._end_canvas_pan()
                    return
            except Exception:
                pass
            pos = args.GetPosition(target)
            dx = float(pos.X) - float(self._pan_last.X)
            dy = float(pos.Y) - float(self._pan_last.Y)
            if abs(dx) > 0.01 or abs(dy) > 0.01:
                self._view_pan_x = float(self._view_pan_x) + dx
                self._view_pan_y = float(self._view_pan_y) + dy
                self._pan_last = pos
                try:
                    self._redraw_canvas(detail=False)
                except Exception:
                    pass
            try:
                args.Handled = True
            except Exception:
                pass
        except Exception:
            pass

    def _on_canvas_pan_up(self, sender, args):
        if not self._panning:
            return
        self._end_canvas_pan()
        try:
            args.Handled = True
        except Exception:
            pass

    def _on_canvas_pan_lost(self, sender, args):
        self._end_canvas_pan()

    def _end_canvas_pan(self):
        was = self._panning
        target = self._nav_capture_target()
        self._panning = False
        self._pan_last = None
        if target is None:
            self._pan_origin_cursor = None
            return
        try:
            if target.IsMouseCaptured:
                target.ReleaseMouseCapture()
        except Exception:
            pass
        try:
            target.Cursor = (
                self._pan_origin_cursor
                if self._pan_origin_cursor is not None
                else Cursors.Arrow
            )
        except Exception:
            try:
                target.Cursor = Cursors.Arrow
            except Exception:
                pass
        self._pan_origin_cursor = None
        if was:
            try:
                self._redraw_canvas(detail=True)
            except Exception:
                pass

    def _on_window_size_changed(self, sender, args):
        self._request_canvas_relayout()

    def _request_canvas_relayout(self):
        """Tras layout / maximizar: ajustar canvas al host y redibujar."""
        self._resize_token += 1
        token = self._resize_token

        def _run():
            if token != self._resize_token:
                return
            try:
                changed = self._sync_canvas_size_from_host()
            except Exception:
                changed = False
            if changed or not self._schema_built:
                try:
                    self._redraw_canvas(detail=True)
                except Exception:
                    pass
                self._schema_built = True

        try:
            self._win.Dispatcher.BeginInvoke(
                DispatcherPriority.Loaded, Action(_run)
            )
        except Exception:
            try:
                _run()
            except Exception:
                pass

    def _sync_canvas_size_from_host(self):
        """
        Hace que CnvBar ocupe el área útil de BdrCanvasHost.

        Returns:
            True si cambió el tamaño de dibujo de forma relevante.
        """
        host = self._canvas_host
        if host is None and self._win is not None:
            try:
                host = self._win.FindName(u"BdrCanvasHost")
                self._canvas_host = host
            except Exception:
                host = None
        w = 0.0
        h = 0.0
        if host is not None:
            try:
                w = float(host.ActualWidth or 0.0)
                h = float(host.ActualHeight or 0.0)
            except Exception:
                w = h = 0.0
        if w < 40.0 or h < 40.0:
            # Fallback: área del panel izquierdo vía ventana − rail.
            try:
                ww = float(self._win.ActualWidth or 0.0)
                wh = float(self._win.ActualHeight or 0.0)
                # Padding 18*2 + márgenes internos ~36 + header/footer ~160 + rail.
                w = max(float(_CANVAS_MIN_W), ww - float(_RAIL_W) - 80.0)
                h = max(float(_CANVAS_MIN_H), wh - 220.0)
            except Exception:
                w = float(_CANVAS_MIN_W)
                h = float(_CANVAS_MIN_H)
        w = max(float(_CANVAS_MIN_W), w)
        h = max(float(_CANVAS_MIN_H), h)
        prev_w = float(self._canvas_w)
        prev_h = float(self._canvas_h)
        # Ignorar micro-cambios (evita redraw en bucle).
        if abs(w - prev_w) < 2.0 and abs(h - prev_h) < 2.0:
            if self._cnv is not None:
                try:
                    if abs(float(self._cnv.Width or 0.0) - w) > 1.0:
                        self._cnv.Width = w
                    if abs(float(self._cnv.Height or 0.0) - h) > 1.0:
                        self._cnv.Height = h
                except Exception:
                    pass
            return False
        self._canvas_w = w
        self._canvas_h = h
        if self._cnv is not None:
            try:
                self._cnv.Width = w
                self._cnv.Height = h
            except Exception:
                pass
        return True

    def _canvas_margin(self):
        """Margen proporcional al tamaño del canvas (mín. 40 px)."""
        try:
            side = min(float(self._canvas_w), float(self._canvas_h))
            return max(float(_MARGIN), min(72.0, side * 0.045))
        except Exception:
            return float(_MARGIN)

    def _invoke_ui(self, action):
        try:
            if not bool(getattr(self._win, u"IsVisible", False)):
                action()
                return
        except Exception:
            pass
        try:
            self._win.Dispatcher.BeginInvoke(
                DispatcherPriority.Normal, Action(action)
            )
        except Exception:
            try:
                action()
            except Exception:
                pass

    def _invoke_ui_sync(self, action):
        """Ejecuta en el Dispatcher WPF y espera (p. ej. tras un Pick en ExternalEvent)."""
        try:
            disp = self._win.Dispatcher
            if disp.CheckAccess():
                action()
                return
            disp.Invoke(DispatcherPriority.Normal, Action(action))
        except Exception:
            try:
                action()
            except Exception:
                pass

    def _ui_set_status(self, msg, err=False, ok=False):
        def _do():
            if self._txt_status is None:
                return
            self._txt_status.Text = _as_unicode(msg)
            if err:
                self._txt_status.Foreground = _brush(u"#E57373")
            elif ok:
                self._txt_status.Foreground = _brush(u"#4ADE80")
            else:
                self._txt_status.Foreground = _brush(FG_MUTED)

        self._invoke_ui(_do)

    def _ui_set_alert(self, msg, kind=u"info"):
        def _do():
            if self._txt_alert is None:
                return
            self._txt_alert.Text = _as_unicode(msg)
            if kind == u"warn":
                self._txt_alert.Foreground = _brush(u"#F5B0B0")
            elif kind == u"ok":
                self._txt_alert.Foreground = _brush(u"#9AE6B4")
            else:
                self._txt_alert.Foreground = _brush(FG_BODY)

        self._invoke_ui(_do)

    def _set_busy(self, busy, status_msg=None):
        self._busy = bool(busy)
        self._ui_enable_actions(not self._busy)
        if status_msg:
            self._ui_set_status(status_msg)

    def _ui_enable_actions(self, enabled):
        def _do():
            can = bool(enabled) and not self._busy
            for btn in (
                self._btn_apply,
                self._btn_limpiar,
                self._btn_close,
            ):
                if btn is not None:
                    btn.IsEnabled = can
            if self._cmb_grade is not None:
                self._cmb_grade.IsEnabled = can
            if self._cmb_lap_mode is not None:
                self._cmb_lap_mode.IsEnabled = can

        self._invoke_ui(_do)

    def _selected_grade(self):
        try:
            item = self._cmb_grade.SelectedItem
            if item is None:
                return None
            content = getattr(item, u"Content", None)
            return grade_from_combo_text(content)
        except Exception:
            return None

    def _update_meta(self):
        lines = [
            u"Id {0} · layout {1} · {2} tramo(s)".format(
                self._session.get(u"rebar_id_int") or u"?",
                self._session.get(u"layout") or u"—",
                self._session.get(u"n_segments") or 0,
            )
        ]
        if self._diameter_mm is not None:
            lines.append(u"ø nominal ≈ {0:.0f} mm".format(float(self._diameter_mm)))
        lines.append(u"Traslape tabla ≈ {0:.0f} mm".format(self._lap_mm))
        mode = self._selected_lap_mode()
        if mode == LAP_MODE_ENDPOINT_PREV:
            lines.append(u"Modo traslape: endpoint anterior (+L)")
        elif mode == LAP_MODE_ENDPOINT_NEXT:
            lines.append(u"Modo traslape: endpoint siguiente (−L)")
        else:
            lines.append(u"Modo traslape: simétrico ± L/2")
        lines.append(u"Longitud centerline ≈ {0:.0f} mm".format(self._total_mm))
        plane = self._session.get(u"plan_plane")
        if plane:
            if _as_unicode(plane) == u"view":
                if self._session.get(u"view_is_section_elevation"):
                    lines.append(
                        u"Esquema: sección / alzado (Right y Up de la vista)"
                    )
                else:
                    lines.append(u"Esquema: proyección de la vista activa")
            elif _as_unicode(plane) == u"normal":
                lines.append(u"Esquema: proyección en el plano de la barra")
            else:
                lines.append(
                    u"Esquema proyectado: plano {0}".format(
                        _as_unicode(plane).upper()
                    )
                )
        if self._context_n_elems > 0:
            pl_n = len(self._context_polylines_uv or [])
            rect_n = len(self._context_fill_rects_uv or [])
            if rect_n > 0:
                lines.append(
                    u"Hormigón en vista: {0} elem. (muro/viga/fund./col./losa), {1} silueta(s)".format(
                        self._context_n_elems, rect_n
                    )
                )
            elif pl_n > 0:
                lines.append(
                    u"Hormigón visible en vista: {0} elemento(s), trazos {1}".format(
                        self._context_n_elems, pl_n
                    )
                )
            else:
                lines.append(
                    u"Hormigón visible en vista: {0} elemento(s) (sin trazos proyectables)".format(
                        self._context_n_elems
                    )
                )
        n_pos = self._session.get(u"n_positions") or 1
        layout_name = self._session.get(u"layout") or u""
        if n_pos > 1 or layout_name == u"MaximumSpacing":
            lines.append(layout_label(layout_name, n_pos))
        if self._txt_meta is not None:
            self._txt_meta.Text = u"\n".join(lines)

    def _context_bbox_uv(self):
        pts = []
        for pl in self._context_polylines_uv or []:
            for p in pl or []:
                pts.append(p)
        for rect in self._context_fill_rects_uv or []:
            try:
                u0, u1, v0, v1 = rect
                pts.append([float(u0), float(v0)])
                pts.append([float(u1), float(v1)])
            except Exception:
                pass
        return pts

    def _draw_concrete_rect_uv(self, rect_uv, mapping):
        if not rect_uv or len(rect_uv) < 4:
            return
        try:
            u0, u1, v0, v1 = rect_uv
            u0, u1, v0, v1 = float(u0), float(u1), float(v0), float(v1)
        except Exception:
            return
        x0, y0 = map_uv_to_canvas_px(u0, v0, mapping)
        x1, y1 = map_uv_to_canvas_px(u1, v1, mapping)
        left = min(float(x0), float(x1))
        top = min(float(y0), float(y1))
        width = max(1.0, abs(float(x1) - float(x0)))
        height = max(1.0, abs(float(y1) - float(y0)))
        rect = Rectangle()
        rect.Width = width
        rect.Height = height
        rect.Fill = _brush_rgb(
            _COLOR_CONCRETE_FILL_RGB[0],
            _COLOR_CONCRETE_FILL_RGB[1],
            _COLOR_CONCRETE_FILL_RGB[2],
            _COLOR_CONCRETE_FILL_A,
        )
        rect.Stroke = _brush(_COLOR_CONCRETE_EDGE, _COLOR_CONCRETE_EDGE_A)
        try:
            rect.StrokeThickness = float(_COLOR_CONCRETE_STROKE)
        except Exception:
            pass
        try:
            rect.SnapsToDevicePixels = False
            rect.UseLayoutRounding = False
        except Exception:
            pass
        Canvas.SetLeft(rect, left)
        Canvas.SetTop(rect, top)
        try:
            Canvas.SetZIndex(rect, 0)
        except Exception:
            pass
        self._cnv.Children.Add(rect)

    def _rebuild_canvas_mapping(self):
        bbox_extra = self._context_bbox_uv()
        all_uv = list(self._plan_uv or [])
        all_uv.extend(bbox_extra)
        self._canvas_mapping = compute_canvas_mapping(
            all_uv,
            float(self._canvas_w),
            float(self._canvas_h),
            self._canvas_margin(),
            swap_uv=False,
            flip_v=self._plan_flip_v,
        )
        z = self._clamp_view_zoom(self._view_zoom)
        try:
            ox = float(self._canvas_mapping.get(u"ox") or 0.0)
            oy = float(self._canvas_mapping.get(u"oy") or 0.0)
            scale = float(self._canvas_mapping.get(u"scale") or 1.0)
        except Exception:
            ox, oy, scale = 0.0, 0.0, 1.0
        cx, cy = self._canvas_center_px()
        self._canvas_mapping[u"scale"] = scale * z
        self._canvas_mapping[u"ox"] = (
            cx + (ox - cx) * z + float(self._view_pan_x or 0.0)
        )
        self._canvas_mapping[u"oy"] = (
            cy + (oy - cy) * z + float(self._view_pan_y or 0.0)
        )
        return self._canvas_mapping

    def _rebuild_plan_px(self):
        mapping = self._rebuild_canvas_mapping()
        self._plan_px = map_polyline_uv_to_canvas_px(self._plan_uv, mapping)
        self._scale = float(mapping.get(u"scale") or 1.0)
        return self._plan_px

    def _style_canvas_stroke(self, shape, thickness):
        """Antialiasing + caps/joins redondos (trazos más suaves en el Canvas)."""
        try:
            shape.StrokeThickness = float(thickness)
            shape.StrokeStartLineCap = PenLineCap.Round
            shape.StrokeEndLineCap = PenLineCap.Round
            shape.StrokeLineJoin = PenLineJoin.Round
            shape.StrokeMiterLimit = 1.0
        except Exception:
            try:
                shape.StrokeThickness = float(thickness)
            except Exception:
                pass
        try:
            shape.SnapsToDevicePixels = False
            shape.UseLayoutRounding = False
        except Exception:
            pass
        try:
            RenderOptions.SetEdgeMode(shape, EdgeMode.Unspecified)
        except Exception:
            pass

    def _add_polyline_stroke(self, points_px, stroke, thickness, opacity=255, soft=False):
        if not points_px or len(points_px) < 2:
            return
        # Halo opcional (desactivado por defecto: engorda el trazo).
        if soft and float(thickness) >= 2.5 and int(opacity) >= 50:
            try:
                halo_a = max(18, min(70, int(opacity * 0.28)))
                self._add_polyline_stroke(
                    points_px,
                    stroke,
                    float(thickness) * 1.8,
                    halo_a,
                    soft=False,
                )
            except Exception:
                pass
        brush = _brush(stroke, opacity)
        try:
            pl = Polyline()
            pts = PointCollection()
            for p in points_px:
                pts.Add(Point(float(p[0]), float(p[1])))
            pl.Points = pts
            pl.Stroke = brush
            self._style_canvas_stroke(pl, thickness)
            self._cnv.Children.Add(pl)
            return
        except Exception:
            pass
        # Fallback: segmentos Line
        for i in range(1, len(points_px)):
            ln = Line()
            ln.X1 = float(points_px[i - 1][0])
            ln.Y1 = float(points_px[i - 1][1])
            ln.X2 = float(points_px[i][0])
            ln.Y2 = float(points_px[i][1])
            ln.Stroke = brush
            self._style_canvas_stroke(ln, thickness)
            self._cnv.Children.Add(ln)

    def _add_canvas_label(self, text, x, y, color_hex, em_size=12.0, bold=True):
        """
        Etiqueta en el Canvas (T# / C#).

        En este host, ``FormattedText``/``TextBlock``/``RenderTargetBitmap`` suelen
        no pintar. Se dibuja con glifos ``Line`` (mismo truco que ya funcionaba).
        """
        if self._cnv is None:
            return False
        s = _as_unicode(text)
        if not s:
            return False
        lx = float(x)
        ly = float(y)
        try:
            h = _as_unicode(color_hex).lstrip(u"#")
            br = SolidColorBrush(
                Color.FromArgb(
                    Byte(255),
                    Byte(int(h[0:2], 16)),
                    Byte(int(h[2:4], 16)),
                    Byte(int(h[4:6], 16)),
                )
            )
        except Exception:
            br = _brush(color_hex)

        try:
            return self._add_canvas_label_strokes(s, lx, ly, br, float(em_size))
        except Exception:
            return False

    def _add_canvas_label_strokes(self, text, x, y, brush, em_size):
        """Dibuja T/C/dígitos con segmentos Line (fallback garantizado)."""
        glyphs = {
            u"T": ((0.0, 0.0, 1.0, 0.0), (0.5, 0.0, 0.5, 1.0)),
            u"C": (
                (0.85, 0.15, 0.5, 0.0),
                (0.5, 0.0, 0.15, 0.15),
                (0.15, 0.15, 0.15, 0.85),
                (0.15, 0.85, 0.5, 1.0),
                (0.5, 1.0, 0.85, 0.85),
            ),
            u"1": ((0.35, 0.2, 0.55, 0.0), (0.55, 0.0, 0.55, 1.0), (0.3, 1.0, 0.8, 1.0)),
            u"2": (
                (0.15, 0.2, 0.5, 0.0),
                (0.5, 0.0, 0.85, 0.2),
                (0.85, 0.2, 0.2, 1.0),
                (0.2, 1.0, 0.9, 1.0),
            ),
            u"3": (
                (0.2, 0.1, 0.8, 0.1),
                (0.8, 0.1, 0.8, 0.45),
                (0.35, 0.5, 0.8, 0.5),
                (0.8, 0.55, 0.8, 0.9),
                (0.8, 0.9, 0.2, 0.9),
            ),
            u"4": (
                (0.7, 0.0, 0.2, 0.65),
                (0.2, 0.65, 0.9, 0.65),
                (0.7, 0.0, 0.7, 1.0),
            ),
            u"5": (
                (0.8, 0.0, 0.25, 0.0),
                (0.25, 0.0, 0.25, 0.45),
                (0.25, 0.45, 0.75, 0.45),
                (0.75, 0.45, 0.75, 1.0),
                (0.75, 1.0, 0.2, 1.0),
            ),
            u"6": (
                (0.75, 0.1, 0.3, 0.1),
                (0.3, 0.1, 0.25, 0.55),
                (0.25, 0.55, 0.7, 0.5),
                (0.7, 0.5, 0.7, 0.95),
                (0.7, 0.95, 0.25, 0.95),
                (0.25, 0.95, 0.25, 0.55),
            ),
            u"7": ((0.15, 0.0, 0.9, 0.0), (0.9, 0.0, 0.35, 1.0)),
            u"8": (
                (0.5, 0.0, 0.2, 0.25),
                (0.2, 0.25, 0.5, 0.5),
                (0.5, 0.5, 0.8, 0.25),
                (0.8, 0.25, 0.5, 0.0),
                (0.5, 0.5, 0.2, 0.75),
                (0.2, 0.75, 0.5, 1.0),
                (0.5, 1.0, 0.8, 0.75),
                (0.8, 0.75, 0.5, 0.5),
            ),
            u"9": (
                (0.75, 0.45, 0.75, 0.1),
                (0.75, 0.1, 0.3, 0.1),
                (0.3, 0.1, 0.3, 0.45),
                (0.3, 0.45, 0.75, 0.45),
                (0.75, 0.45, 0.75, 0.9),
                (0.75, 0.9, 0.3, 0.9),
            ),
            u"0": (
                (0.3, 0.15, 0.7, 0.15),
                (0.7, 0.15, 0.7, 0.85),
                (0.7, 0.85, 0.3, 0.85),
                (0.3, 0.85, 0.3, 0.15),
            ),
        }
        h = max(12.0, float(em_size))
        w = h * 0.62
        gap = h * 0.18
        ok = False
        cursor = float(x)
        for ch in _as_unicode(text):
            key = ch.upper() if ch.isalpha() else ch
            segs = glyphs.get(key)
            if not segs:
                cursor += w * 0.5
                continue
            for x0, y0, x1, y1 in segs:
                ln = Line()
                ln.X1 = cursor + float(x0) * w
                ln.Y1 = float(y) + float(y0) * h
                ln.X2 = cursor + float(x1) * w
                ln.Y2 = float(y) + float(y1) * h
                ln.Stroke = brush
                ln.StrokeThickness = max(1.5, h * 0.12)
                try:
                    ln.StrokeStartLineCap = PenLineCap.Round
                    ln.StrokeEndLineCap = PenLineCap.Round
                except Exception:
                    pass
                self._cnv.Children.Add(ln)
                ok = True
            cursor += w + gap
        return ok

    def _mm_to_canvas(self, mm):
        if not self._plan_px:
            self._rebuild_plan_px()
        pt_px = point_at_arc_length_uv(self._plan_px, self._plan_arc_mm, mm)
        if pt_px is None:
            return self._canvas_margin(), self._canvas_margin()
        return float(pt_px[0]), float(pt_px[1])

    def _nearest_cut_index_at_px(self, x, y):
        if not self._cuts_mm:
            return -1
        if not self._plan_px:
            self._rebuild_plan_px()
        best_i = -1
        best_d = None
        for i, c in enumerate(sorted(self._cuts_mm)):
            cx, cy = self._mm_to_canvas(c)
            d = ((float(x) - cx) ** 2 + (float(y) - cy) ** 2) ** 0.5
            if d <= _CUT_HIT_PX and (best_d is None or d < best_d):
                best_d = d
                best_i = i
        return best_i

    def _redraw_canvas(self, detail=True):
        if self._cnv is None:
            return
        self._cnv.Children.Clear()
        pts = self._rebuild_plan_px()
        if len(pts) < 2:
            return

        # Hormigón visible (sección/alzado: siluetas AABB como Armado Vigas).
        mapping = self._canvas_mapping or self._rebuild_canvas_mapping()
        for rect_uv in self._context_fill_rects_uv or []:
            self._draw_concrete_rect_uv(rect_uv, mapping)
        if not self._context_fill_rects_uv:
            for pl_uv in self._context_polylines_uv or []:
                if not pl_uv or len(pl_uv) < 2:
                    continue
                ctx_px = map_polyline_uv_to_canvas_px(pl_uv, mapping)
                if len(ctx_px) >= 2:
                    self._add_polyline_stroke(
                        ctx_px,
                        _COLOR_CONCRETE_EDGE,
                        _COLOR_CONCRETE_STROKE + 0.25,
                        _COLOR_CONCRETE_EDGE_A,
                        soft=False,
                    )

        # Base bar.
        self._add_polyline_stroke(pts, _COLOR_BAR, _BAR_THICK, 255)

        mode = self._selected_lap_mode()
        spans = build_spans_mm(self._total_mm, self._cuts_mm) if self._cuts_mm else []
        fab_intervals = []
        if self._cuts_mm:
            try:
                fab_intervals = piece_intervals_with_lap(
                    self._total_mm, self._cuts_mm, self._lap_mm, lap_mode=mode
                )
            except Exception:
                fab_intervals = []

        # Tramos fabricados (se solapan en la zona de empalme).
        draw_intervals = fab_intervals
        if not draw_intervals and spans:
            draw_intervals = [
                (
                    float(s.get(u"start_mm") or 0.0),
                    float(s.get(u"end_mm") or 0.0),
                )
                for s in spans
            ]
        for i, (a_mm, b_mm) in enumerate(draw_intervals):
            if float(b_mm) - float(a_mm) < 1e-6:
                continue
            active = int(self._active_tramo_index) == i
            n_samp = max(4, int((b_mm - a_mm) / max(1.0, _SPAN_SAMPLE_MM)) + 1)
            if not detail:
                n_samp = max(3, n_samp // 2)
            sub_px = []
            for k in range(n_samp + 1):
                s = a_mm + (b_mm - a_mm) * (float(k) / float(n_samp))
                p = point_at_arc_length_uv(self._plan_px, self._plan_arc_mm, s)
                if p is not None:
                    sub_px.append(p)
            color = _tramo_color(i)
            alpha = 210 if active else 150
            thick = _BAR_THICK + (1.2 if active else 0.4)
            self._add_polyline_stroke(sub_px, color, thick, alpha)

        # Empalme encima del solape (zona según modo).
        sorted_cuts = sorted(self._cuts_mm)
        for c in sorted_cuts:
            a_mm, b_mm = lap_zone_around_cut(c, self._lap_mm, mode)
            a_mm = max(0.0, min(self._total_mm, a_mm))
            b_mm = max(0.0, min(self._total_mm, b_mm))
            if b_mm - a_mm < 1e-6:
                continue
            n_samp = 10 if detail else 5
            sub_px = []
            for k in range(n_samp + 1):
                s = a_mm + (b_mm - a_mm) * (float(k) / float(n_samp))
                p = point_at_arc_length_uv(self._plan_px, self._plan_arc_mm, s)
                if p is not None:
                    sub_px.append(p)
            self._add_polyline_stroke(sub_px, _COLOR_LAP, _BAR_THICK + 3.0, 90)

        for i, c in enumerate(sorted_cuts):
            try:
                cx, cy = self._mm_to_canvas(c)
                tx, ty = tangent_at_arc_length_uv(
                    self._plan_px, self._plan_arc_mm, c
                )
                nx, ny = -ty, tx
                cut_ln = Line()
                cut_ln.X1 = cx - nx * _CUT_MARK_PX
                cut_ln.Y1 = cy - ny * _CUT_MARK_PX
                cut_ln.X2 = cx + nx * _CUT_MARK_PX
                cut_ln.Y2 = cy + ny * _CUT_MARK_PX
                cut_ln.Stroke = _brush(_COLOR_CUT, 220)
                self._style_canvas_stroke(cut_ln, 1.5)
                self._cnv.Children.Add(cut_ln)
                self._add_canvas_label(
                    u"C{0}".format(i + 1),
                    cx + nx * (_CUT_MARK_PX + 4.0) - 8.0,
                    cy + ny * (_CUT_MARK_PX + 4.0) - 8.0,
                    _COLOR_CUT,
                    10.0,
                )
            except Exception:
                continue

        # Etiquetas T1, T2… junto a cada tramo (mismo índice/color que configuradores).
        try:
            self._draw_tramo_labels_on_canvas(spans)
        except Exception:
            pass

    def _draw_tramo_labels_on_canvas(self, spans):
        """Dibuja T1, T2… en el punto medio de cada tramo del esquema."""
        if self._cnv is None or not spans:
            return
        for i, span in enumerate(spans):
            try:
                mid = span_midpoint_mm(span)
                cx, cy = self._mm_to_canvas(mid)
                nx, ny = 1.0, 0.0
                try:
                    tan = tangent_at_arc_length_uv(
                        self._plan_px, self._plan_arc_mm, mid
                    )
                    if tan is not None:
                        tx = float(tan[0])
                        ty = float(tan[1])
                        qx, qy = -ty, tx
                        qlen = (qx * qx + qy * qy) ** 0.5
                        if qlen > 1e-9:
                            qx, qy = qx / qlen, qy / qlen
                            if qx < 0:
                                qx, qy = -qx, -qy
                            if abs(qx) < 0.35:
                                qx, qy = 1.0, 0.0
                            nx, ny = qx, qy
                except Exception:
                    nx, ny = 1.0, 0.0

                active = int(self._active_tramo_index) == i
                color = _tramo_color(i)
                offset = 28.0 if active else 24.0
                lx = float(cx) + float(nx) * offset
                ly = float(cy) + float(ny) * offset - 10.0
                if (
                    lx < -48.0
                    or ly < -32.0
                    or lx > float(self._canvas_w) + 48.0
                    or ly > float(self._canvas_h) + 32.0
                ):
                    continue
                lx = max(4.0, min(float(self._canvas_w) - 44.0, lx))
                ly = max(4.0, min(float(self._canvas_h) - 24.0, ly))
                em = 13.0 if active else 11.0
                self._add_canvas_label(
                    tramo_label(i),
                    lx + 1.0,
                    ly + 1.0,
                    u"#020617",
                    em,
                    bold=True,
                )
                self._add_canvas_label(
                    tramo_label(i), lx, ly, color, em, bold=True
                )
            except Exception:
                continue

    def _fabricated_lengths_mm(self):
        if not self._cuts_mm:
            return []
        return fabricated_lengths_mm(
            self._total_mm,
            self._cuts_mm,
            self._lap_mm,
            lap_mode=self._selected_lap_mode(),
        )

    def _ensure_tramos_panel(self):
        """
        Garantiza un StackPanel vivo para los editores de largo.

        No depende solo de FindName(x:Name en StackPanel): con XamlReader +
        caché AppDomain a veces el nombre no resuelve y el panel queda vacío
        aunque el canvas sí pinte los cortes.
        """
        if self._pnl_tramos is not None:
            return True
        host = self._brd_tramos_host
        if host is None and self._win is not None:
            try:
                host = self._win.FindName(u"BrdTramosHost")
                self._brd_tramos_host = host
            except Exception:
                host = None
        # Compat: XAML antiguo con StackPanel x:Name="PnlTramos"
        if host is None and self._win is not None:
            try:
                legacy = self._win.FindName(u"PnlTramos")
            except Exception:
                legacy = None
            if legacy is not None and hasattr(legacy, u"Children"):
                self._pnl_tramos = legacy
                return True
        if host is None:
            return False
        pnl = StackPanel()
        pnl.Margin = Thickness(0, 0, 0, 0)
        try:
            host.Child = pnl
        except Exception:
            try:
                if hasattr(host, u"Children"):
                    host.Children.Clear()
                    host.Children.Add(pnl)
                else:
                    return False
            except Exception:
                return False
        self._pnl_tramos = pnl
        return True

    def _make_tramo_editor_block(self, span_index, span, fab_len):
        block = StackPanel()
        block.Margin = Thickness(0, 0, 0, 8)
        block.Tag = span_index
        try:
            block.MouseEnter += MouseEventHandler(self._on_tramo_mouse_enter)
            block.MouseLeave += MouseEventHandler(self._on_tramo_mouse_leave)
        except Exception:
            pass

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(0, 0, 0, 2)

        swatch = TextBlock()
        swatch.Text = u"●"
        swatch.Foreground = _brush(_tramo_color(span_index))
        swatch.FontSize = 12.0
        swatch.VerticalAlignment = VerticalAlignment.Center
        swatch.Margin = Thickness(0, 0, 4, 0)
        row.Children.Add(swatch)

        lbl = TextBlock()
        lbl.Text = tramo_label(span_index)
        lbl.Foreground = _brush(_tramo_color(span_index))
        lbl.FontSize = 11.0
        lbl.FontWeight = FontWeights.SemiBold
        lbl.VerticalAlignment = VerticalAlignment.Center
        lbl.Width = 32.0
        row.Children.Add(lbl)

        display_len = fab_len
        if display_len is None:
            display_len = float(span.get(u"length_mm") or 0)
        display_len = ceil_mm_to_nearest_10(display_len)
        tb = TextBox()
        tb.Text = u"{0:.0f}".format(display_len)
        tb.Width = 72.0
        tb.Height = 26.0
        tb.VerticalContentAlignment = VerticalAlignment.Center
        tb.HorizontalContentAlignment = HorizontalAlignment.Right
        tb.Tag = span_index
        try:
            tb.Style = self._win.FindResource(u"BimToolsTextBoxDark")
        except Exception:
            pass
        try:
            tb.GotFocus += RoutedEventHandler(self._on_tramo_got_focus)
        except Exception:
            pass
        try:
            tb.LostFocus += RoutedEventHandler(self._on_span_lost_focus)
        except Exception:
            pass
        try:
            tb.KeyDown += KeyEventHandler(self._on_span_key_down)
        except Exception:
            pass
        row.Children.Add(tb)
        self._tramo_editors[int(span_index)] = tb

        mm_lbl = TextBlock()
        mm_lbl.Text = u" mm"
        mm_lbl.Foreground = _brush(FG_BODY)
        mm_lbl.VerticalAlignment = VerticalAlignment.Center
        mm_lbl.Margin = Thickness(4, 0, 0, 0)
        row.Children.Add(mm_lbl)

        block.Children.Add(row)

        nominal = float(span.get(u"length_mm") or 0)
        if fab_len is not None and abs(float(fab_len) - nominal) > 0.5:
            sub = TextBlock()
            sub.Text = u"Centerline ≈ {0:.0f} mm".format(nominal)
            sub.Foreground = _brush(FG_MUTED)
            sub.FontSize = 9.0
            block.Children.Add(sub)

        return block

    def _rebuild_tramos_panel(self):
        """Editor de largo fabricado por tramo + preview con traslape."""
        if not self._ensure_tramos_panel():
            self._ui_set_status(
                u"No se pudo montar el panel de largos de tramo (UI).",
                err=True,
            )
            return

        self._tramo_editors = {}
        blocks = []
        try:
            if not self._cuts_mm:
                self._active_tramo_index = -1
                hint = TextBlock()
                hint.Text = (
                    u"Aún no hay cortes. Clic en el esquema de alzado (cerca de "
                    u"la barra) para añadir cortes. Aquí aparecerán T1, T2… "
                    u"para afinar cada largo."
                )
                hint.TextWrapping = TextWrapping.Wrap
                hint.Foreground = _brush(FG_BODY)
                hint.FontSize = float(FONT_SIZE_HINT)
                blocks.append(hint)
            else:
                spans = build_spans_mm(self._total_mm, self._cuts_mm)
                fab = self._fabricated_lengths_mm()
                for i, span in enumerate(spans):
                    fab_len = fab[i] if i < len(fab) else None
                    blocks.append(self._make_tramo_editor_block(i, span, fab_len))
                if self._active_tramo_index >= len(spans):
                    self._active_tramo_index = -1
        except Exception as ex:
            err = TextBlock()
            err.Text = u"Error al crear editores: {0}".format(_as_unicode(ex))
            err.TextWrapping = TextWrapping.Wrap
            err.Foreground = _brush(u"#E57373")
            err.FontSize = 11.0
            blocks = [err]

        self._pnl_tramos.Children.Clear()
        for block in blocks:
            self._pnl_tramos.Children.Add(block)

    def _refresh_all(self, update_status=True, detail=True):
        if self._rebuilding:
            return
        self._rebuilding = True
        try:
            self._redraw_canvas(detail=detail)
            try:
                self._rebuild_tramos_panel()
            except Exception as ex:
                self._ui_set_status(
                    u"Error al montar tramos: {0}".format(_as_unicode(ex)), err=True
                )
            if update_status:
                n = len(self._cuts_mm)
                self._ui_set_status(
                    u"{0} corte(s) → {1} tramo(s) fabricado(s)".format(
                        n, n + 1 if n else 0
                    )
                )
        finally:
            self._rebuilding = False

    def _build_schema_ready(self):
        """
        Monta canvas + paneles; el tamaño final se aplica tras layout/maximizar
        vía ``_request_canvas_relayout``.
        """
        if self._schema_built:
            return
        try:
            self._sync_canvas_size_from_host()
        except Exception:
            pass
        try:
            self._refresh_all(update_status=False, detail=True)
            self._ui_set_status(
                u"Barra Id {0} · {1:.0f} mm · traslape {2:.0f} mm".format(
                    self._session.get(u"rebar_id_int") or u"?",
                    self._total_mm,
                    self._lap_mm,
                ),
                ok=True,
            )
            self._schema_built = True
        except Exception as ex:
            self._ui_set_status(_as_unicode(ex), err=True)

    def _validate_and_set_cuts(self, cuts, show_err=True):
        ok, msg, sorted_cuts = validate_cuts_with_lap(
            self._total_mm,
            cuts,
            self._lap_mm,
            _MIN_PIECE_MM,
            lap_mode=self._selected_lap_mode(),
        )
        if not ok:
            if show_err:
                self._ui_set_status(msg, err=True)
                self._ui_set_alert(msg, u"warn")
            return False
        self._cuts_mm = list(sorted_cuts)
        self._refresh_all()
        return True

    def _try_add_cut(self, mm_val):
        trial = sorted(self._cuts_mm + [float(mm_val)])
        return self._validate_and_set_cuts(trial)

    def _remove_cut_at_index(self, index):
        sorted_cuts = sorted(self._cuts_mm)
        if index < 0 or index >= len(sorted_cuts):
            return
        val = sorted_cuts[index]
        self._cuts_mm = [c for c in self._cuts_mm if abs(c - val) > 1e-6]
        self._refresh_all()

    def _commit_span(self, span_index, text):
        try:
            new_len = float(_as_unicode(text).strip().replace(u",", u"."))
        except Exception:
            self._ui_set_status(u"Longitud de tramo inválida.", err=True)
            self._refresh_all()
            return
        new_len = ceil_mm_to_nearest_10(new_len)
        ok, msg, new_cuts = set_span_fabricated_length_mm(
            self._total_mm,
            self._cuts_mm,
            int(span_index),
            new_len,
            self._lap_mm,
            _MIN_PIECE_MM,
            lap_mode=self._selected_lap_mode(),
        )
        if not ok:
            self._ui_set_status(msg, err=True)
            self._refresh_all()
            return
        self._cuts_mm = list(new_cuts)
        self._refresh_all()
        self._ui_set_status(
            u"Tramo T{0} → {1:.0f} mm (fabricado)".format(span_index + 1, new_len),
            ok=True,
        )

    def _set_active_tramo(self, span_index, focus_editor=False, redraw=True):
        """
        Relaciona el editor T# con el tramo dibujado en el canvas.

        ``span_index`` -1 limpia el resaltado.
        """
        try:
            idx = int(span_index)
        except Exception:
            idx = -1
        if idx < 0:
            idx = -1
        changed = idx != int(self._active_tramo_index)
        self._active_tramo_index = idx
        if redraw and changed and not self._rebuilding:
            try:
                self._redraw_canvas(detail=True)
            except Exception:
                pass
        if focus_editor and idx >= 0:
            tb = self._tramo_editors.get(idx)
            if tb is not None:
                try:
                    tb.Focus()
                    tb.SelectAll()
                except Exception:
                    try:
                        tb.Focus()
                    except Exception:
                        pass

    def _nearest_mm_at_px(self, x, y):
        """Proyecta un clic del canvas a mm a lo largo de la centerline."""
        if not self._plan_px or not self._plan_arc_mm:
            self._rebuild_plan_px()
        if len(self._plan_px) < 2 or len(self._plan_arc_mm) < 2:
            return None
        best_d = None
        best_mm = None
        step = max(20.0, float(self._total_mm) / 200.0)
        t = 0.0
        while t <= self._total_mm + 1e-6:
            cx, cy = self._mm_to_canvas(t)
            d = ((float(x) - cx) ** 2 + (float(y) - cy) ** 2) ** 0.5
            if best_d is None or d < best_d:
                best_d = d
                best_mm = t
            t += step
        if best_d is None or best_d > _TRAMO_HIT_PX:
            return None
        return best_mm

    def _span_index_at_px(self, x, y):
        if not self._cuts_mm:
            return -1
        mm = self._nearest_mm_at_px(x, y)
        if mm is None:
            return -1
        return span_index_at_mm(self._total_mm, self._cuts_mm, mm)

    def _on_tramo_got_focus(self, sender, args):
        if self._rebuilding:
            return
        try:
            idx = int(getattr(sender, u"Tag", -1))
        except Exception:
            idx = -1
        self._set_active_tramo(idx)

    def _on_tramo_mouse_enter(self, sender, args):
        if self._rebuilding or self._busy:
            return
        try:
            idx = int(getattr(sender, u"Tag", -1))
        except Exception:
            idx = -1
        self._set_active_tramo(idx)

    def _on_tramo_mouse_leave(self, sender, args):
        if self._rebuilding or self._busy:
            return
        # Si un TextBox de tramo tiene foco, conservar su resaltado.
        try:
            focused = getattr(self._win, u"IsActive", True)
            for idx, tb in (self._tramo_editors or {}).items():
                if tb is not None and bool(getattr(tb, u"IsKeyboardFocused", False)):
                    self._set_active_tramo(int(idx))
                    return
            if not focused:
                return
        except Exception:
            pass
        self._set_active_tramo(-1)

    def _on_span_lost_focus(self, sender, args):
        if self._rebuilding:
            return
        tb = sender
        idx = int(getattr(tb, u"Tag", -1))
        self._commit_span(idx, tb.Text)
        self._set_active_tramo(-1)

    def _on_span_key_down(self, sender, args):
        if args.Key == Key.Enter:
            tb = sender
            idx = int(getattr(tb, u"Tag", -1))
            self._commit_span(idx, tb.Text)
            try:
                args.Handled = True
            except Exception:
                pass

    def _on_canvas_click(self, sender, args):
        if self._busy or self._panning:
            return
        pos = args.GetPosition(self._cnv)
        hit = self._nearest_cut_index_at_px(pos.X, pos.Y)
        if hit >= 0:
            self._remove_cut_at_index(hit)
            self._ui_set_status(u"Corte eliminado.")
            return
        if not self._plan_px:
            self._rebuild_plan_px()
        s_mm, dist_px, _px, _py = nearest_arc_length_px(
            self._plan_px, self._plan_arc_mm, pos.X, pos.Y
        )
        if dist_px <= _CUT_ADD_HIT_PX:
            if self._try_add_cut(s_mm):
                self._ui_set_status(
                    u"Corte en {0:.0f} mm".format(float(s_mm)),
                    ok=True,
                )
            return
        if self._cuts_mm:
            span_i = self._span_index_at_px(pos.X, pos.Y)
            if span_i >= 0:
                self._set_active_tramo(span_i, focus_editor=True)
                self._ui_set_status(
                    u"{0} seleccionado en esquema · afine el largo a la derecha.".format(
                        tramo_label(span_i)
                    ),
                    ok=True,
                )
                return
        self._ui_set_status(
            u"Clic más cerca de la barra del esquema para añadir un corte.",
            err=True,
        )

    def _on_limpiar(self, sender, args):
        self._cuts_mm = []
        self._refresh_all()
        self._ui_set_status(u"Cortes limpiados.")
        self._ui_set_alert(
            u"Clic en el esquema de alzado para añadir cortes de nuevo. "
            u"Rueda: zoom · botón medio: desplazar.",
            u"info",
        )

    def _on_lap_mode_changed(self, sender, args):
        if self._busy:
            return
        self._lap_mode = self._selected_lap_mode()
        self._update_lap_mode_hint()
        self._update_meta()
        if self._cuts_mm:
            ok, msg, sorted_cuts = validate_cuts_with_lap(
                self._total_mm,
                self._cuts_mm,
                self._lap_mm,
                _MIN_PIECE_MM,
                lap_mode=self._lap_mode,
            )
            if not ok:
                self._ui_set_status(
                    u"Modo de traslape: revise cortes — {0}".format(msg),
                    err=True,
                )
                self._ui_set_alert(msg, u"warn")
            else:
                self._cuts_mm = list(sorted_cuts)
                self._ui_set_status(
                    u"Modo de traslape actualizado.",
                    ok=True,
                )
        self._refresh_all()

    def _on_grade_changed(self, sender, args):
        if self._busy or self._rebar_id is None:
            return
        try:
            uidoc = self._uiapp.ActiveUIDocument
            if uidoc is None:
                return
            doc = uidoc.Document
            rebar = doc.GetElement(self._rebar_id)
            if not isinstance(rebar, Rebar):
                return
            grade = self._selected_grade()
            d_mm, lap_mm = lap_mm_for_rebar(doc, rebar, grade)
            if lap_mm is None or lap_mm <= 0:
                self._ui_set_status(u"No se pudo calcular traslape.", err=True)
                return
            self._lap_mm = float(lap_mm)
            if d_mm is not None:
                self._diameter_mm = d_mm
            self._update_meta()
            if self._cuts_mm:
                ok, msg, sorted_cuts = validate_cuts_with_lap(
                    self._total_mm,
                    self._cuts_mm,
                    self._lap_mm,
                    _MIN_PIECE_MM,
                    lap_mode=self._selected_lap_mode(),
                )
                if not ok:
                    self._ui_set_status(
                        u"Traslape actualizado; revise cortes: {0}".format(msg),
                        err=True,
                    )
                    self._ui_set_alert(msg, u"warn")
                else:
                    self._cuts_mm = list(sorted_cuts)
                    self._ui_set_status(
                        u"Traslape ≈ {0:.0f} mm (grado actualizado)".format(
                            self._lap_mm
                        ),
                        ok=True,
                    )
            else:
                self._ui_set_status(
                    u"Traslape ≈ {0:.0f} mm (grado actualizado)".format(self._lap_mm),
                    ok=True,
                )
            self._refresh_all()
        except Exception as ex:
            self._ui_set_status(_as_unicode(ex), err=True)

    def _on_apply(self, sender, args):
        global _APPLY_HOLD
        if self._busy or self._apply_started:
            return
        if not self._cuts_mm:
            self._ui_set_status(
                u"Añada al menos un corte (clic en el esquema de alzado).",
                err=True,
            )
            return
        ok, msg, _sorted = validate_cuts_with_lap(
            self._total_mm,
            self._cuts_mm,
            self._lap_mm,
            _MIN_PIECE_MM,
            lap_mode=self._selected_lap_mode(),
        )
        if not ok:
            self._ui_set_status(msg, err=True)
            return

        # Cerrar UI primero; el ExternalEvent ejecuta la división en idle de Revit.
        self._busy = True
        self._apply_started = True
        self._closing_for_apply = True
        # Snapshot: tras Close los controles WPF ya no son fiables.
        self._apply_cuts = sorted(self._cuts_mm)
        self._apply_grade = self._selected_grade()
        self._apply_lap_mode = self._selected_lap_mode()
        _APPLY_HOLD = self
        try:
            if self._win is not None:
                self._win.Close()
        except Exception:
            try:
                if self._win is not None:
                    self._win.Hide()
            except Exception:
                pass
        _clear_singleton()
        try:
            self._ext_apply.Raise()
        except Exception as ex:
            _APPLY_HOLD = None
            self._closing_for_apply = False
            self._apply_started = False
            self._busy = False
            mostrar_aviso(
                self._uiapp,
                u"No se pudo iniciar la división.",
                _as_unicode(ex),
            )

    def apply_in_revit(self, uiapp):
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            mostrar_aviso(uiapp, u"No se pudo dividir.", u"No hay documento activo.")
            return
        doc = uidoc.Document
        rebar = doc.GetElement(self._rebar_id)
        if not isinstance(rebar, Rebar):
            mostrar_aviso(
                uiapp,
                u"No se pudo dividir.",
                u"La barra ya no existe en el modelo.",
            )
            return
        target_view = resolve_active_model_view(uidoc)
        view_for_unobscured = target_view
        if target_view is not None:
            try:
                view_for_unobscured = doc.GetElement(target_view.Id)
            except Exception:
                view_for_unobscured = resolve_active_model_view(uidoc)

        cuts = list(getattr(self, u"_apply_cuts", None) or sorted(self._cuts_mm))
        grade = getattr(self, u"_apply_grade", None)
        if grade is None:
            try:
                grade = self._selected_grade()
            except Exception:
                grade = None
        lap_mode = getattr(self, u"_apply_lap_mode", None)
        if lap_mode is None:
            try:
                lap_mode = self._selected_lap_mode()
            except Exception:
                lap_mode = LAP_MODE_SYMMETRIC

        ok, msg, ids, meta = (False, u"", None, None)
        try:
            with DividirRebarProgress() as progress:
                ok, msg, ids, meta = divide_rebar_at_cuts(
                    doc,
                    rebar,
                    cuts,
                    grade,
                    view=view_for_unobscured,
                    lap_mode=lap_mode,
                    progress=progress,
                )
        except Exception as ex:
            ok = False
            msg = _as_unicode(ex)
        if ok:
            try:
                _clear_selection(uidoc)
            except Exception:
                pass
            return

        mostrar_aviso(uiapp, u"No se pudo dividir.", msg or u"")

    def _finish_after_apply(self):
        """Limpieza tras ExternalEvent (la ventana ya se cerró al aplicar)."""
        global _APPLY_HOLD
        _APPLY_HOLD = None
        self._closing_for_apply = False
        self._apply_started = False
        self._busy = False
        try:
            if self._win is not None and bool(getattr(self._win, u"IsLoaded", False)):
                self._win.Close()
        except Exception:
            pass
        _clear_singleton()

    def _on_manual(self, sender, args):
        _open_manual(self._uiapp)

    def _on_close(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass

    def _on_closed(self, sender, args):
        # Si cerramos para aplicar, el hold mantiene el controlador hasta el idle.
        if self._closing_for_apply:
            return
        _clear_singleton()

    def show(self):
        """Preposición en monitor Revit + maximizar (misma vía que Armado Vigas)."""
        _prepare_window(self._win, self._uiapp)
        self._build_schema_ready()
        try:
            self._win.WindowState = WindowState.Maximized
        except Exception:
            pass
        self._win.Show()
        _maximize_window(self._win)
        _bring_window_to_front(self._win)
        # Tras maximizar, ActualWidth/Height del host ya son válidos.
        self._request_canvas_relayout()
        try:
            self._win.Dispatcher.BeginInvoke(
                DispatcherPriority.ApplicationIdle,
                Action(self._request_canvas_relayout),
            )
        except Exception:
            pass


def show_dividir_rebar_punto_window(revit):
    """Pick de Rebar primero; con barra elegida abre la UI (instancia única)."""
    if _focus_existing(revit):
        return
    uidoc = None
    try:
        uidoc = revit.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        mostrar_aviso(revit, u"No hay documento activo.")
        return
    rebar, err = pick_rebar(uidoc)
    if rebar is None:
        msg = _as_unicode(err or u"")
        low = msg.lower()
        if msg and u"cancel" not in low:
            mostrar_aviso(revit, msg)
        return
    # Solo se necesita el Id/datos de la barra: no dejarla seleccionada
    # (evita cinta Modify de Structural Rebar).
    _clear_selection(uidoc)
    ok, prep_err, session = prepare_division_session(
        uidoc.Document,
        rebar,
        None,
        view=resolve_active_model_view(uidoc),
    )
    if not ok:
        mostrar_aviso(revit, prep_err or u"Barra no elegible.")
        return
    win = DividirRebarPuntoWindow(revit, rebar, session)
    _set_singleton(win)
    win.show()
