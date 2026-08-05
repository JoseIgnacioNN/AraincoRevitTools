# -*- coding: utf-8 -*-
"""UI WPF — Dividir rebar multipunto con traslape por diámetro."""

from __future__ import print_function

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
    ComboBoxItem,
    Orientation,
    SelectionChangedEventHandler,
    StackPanel,
    TextBlock,
    TextBox,
)
from System.Windows.Input import Key, KeyEventHandler, MouseButtonEventHandler, MouseEventHandler
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
from System.Windows.Shapes import Line, Polyline
from System.Windows.Threading import DispatcherPriority
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

from bimtools_ui_tokens import (
    ACCENT_PRIMARY,
    BG_APP,
    BG_INPUT,
    BG_PANEL,
    BORDER,
    FONT_FAMILY,
    FONT_SIZE_BASE,
    FONT_SIZE_HINT,
    FONT_SIZE_STATUS,
    FONT_SIZE_SUBTITLE,
    FONT_SIZE_TITLE,
    FONT_WEIGHT_TITLE,
    FG_BODY,
    FG_MUTED,
    FG_TITLE,
    PAD_PANEL,
    PAD_WINDOW,
    WINDOW_CHROME_TITLE,
)
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from dividir_rebar_punto import _TITULO, _as_unicode, grade_from_combo_text, mostrar_aviso
from dividir_rebar_punto_core import (
    cut_mm_from_pick,
    divide_rebar_at_cuts,
    lap_mm_for_rebar,
    layout_label,
    pick_cut_points_on_rebar,
    pick_rebar,
    prepare_division_session,
    resolve_active_model_view,
)
from dividir_rebar_punto_geom import (
    build_spans_mm,
    fit_polyline_to_canvas,
    lap_zone_around_cut,
    normalize_lap_mode,
    piece_intervals_with_lap,
    point_at_arc_length_uv,
    set_span_length_mm,
    span_index_at_mm,
    span_midpoint_mm,
    tangent_at_arc_length_uv,
    tramo_label,
    validate_cuts_with_lap,
    LAP_MODE_ENDPOINT_NEXT,
    LAP_MODE_ENDPOINT_PREV,
    LAP_MODE_SYMMETRIC,
)
from revit_wpf_window_position import (
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)

_SINGLETON_KEY = u"Arainco_DividirRebarPuntoTraslape_UI"
_MIN_PIECE_MM = 100.0

_CANVAS_W = 420.0
_CANVAS_H = 480.0
_MARGIN = 40.0
# Holgura alrededor del canvas (títulos del esquema, paddings panel/borde).
_SCHEMA_CHROME_H = 72.0
_SCHEMA_CHROME_W = float(PAD_PANEL) * 2.0 + 16.0
_RIGHT_PANEL_MIN_W = 400.0
_LEFT_COL_W = int(_CANVAS_W + _SCHEMA_CHROME_W + 8.0)
# Altura del bloque esquema (= panel izquierdo); el scroll derecho no la supera.
_SCHEMA_BLOCK_H = int(_CANVAS_H + _SCHEMA_CHROME_H)
_WINDOW_W = int(
    float(PAD_WINDOW) * 2.0
    + float(_LEFT_COL_W)
    + 12.0
    + _RIGHT_PANEL_MIN_W
    + 16.0
)
# Altura inicial aproximada (SizeToContent=Height ajusta al contenido real).
_WINDOW_H = int(
    36.0  # cinta Windows
    + float(PAD_WINDOW) * 2.0
    + 78.0  # título + subtítulo
    + float(_SCHEMA_BLOCK_H)
    + 52.0  # footer
    + 8.0
)
_WINDOW_MIN_W = max(880, _WINDOW_W - 60)
_WINDOW_MIN_H = max(520, int(_WINDOW_H * 0.85))

_BAR_THICK = 2.0
_CUT_HIT_PX = 10.0
_CUT_MARK_PX = 14.0

_COLOR_BAR = ACCENT_PRIMARY
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
_XAML_AD_KEY = u"Arainco_DividirRebarXamlCache_v11_lap_mode"
# Claves viejas: forzar XAML fresco tras cambios de panel de tramos.
_XAML_AD_KEYS_PURGE = (
    u"Arainco_DividirRebarXamlCache_v4_tight_height",
    u"Arainco_DividirRebarXamlCache_v5_no_cuts_edit",
    u"Arainco_DividirRebarXamlCache_v6_tramos_edit",
    u"Arainco_DividirRebarXamlCache_v7_tramos_host",
    u"Arainco_DividirRebarXamlCache_v8_pick_preview",
    u"Arainco_DividirRebarXamlCache_v9_pickobjects_finish",
    u"Arainco_DividirRebarXamlCache_v10_tramo_link",
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


def _escape_xaml(text):
    s = _as_unicode(text)
    return (
        s.replace(u"&", u"&amp;")
        .replace(u"<", u"&lt;")
        .replace(u">", u"&gt;")
        .replace(u'"', u"&quot;")
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
    title_esc = _escape_xaml(_TITULO)
    _XAML_CACHE = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="{chrome_title}"
    Width="{window_w}" Height="{window_h}"
    MinWidth="{window_min_w}" MinHeight="{window_min_h}"
    SizeToContent="Height"
    ResizeMode="CanResize"
    WindowStartupLocation="Manual"
    Background="{bg_app}"
    FontFamily="{font_family}"
    FontSize="{font_base}"
    ShowInTaskbar="False">
  <Window.Resources>
{styles}
  </Window.Resources>
  <Border Background="{bg_app}" BorderBrush="{border}" BorderThickness="1"
          Padding="{pad_window}">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,10">
        <TextBlock x:Name="TxtTitle" Text="{title_esc}"
                   Foreground="{fg_title}" FontSize="{font_title}"
                   FontWeight="{font_weight_title}"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0"
                   Foreground="{fg_body}" FontSize="{font_subtitle}"
                   TextWrapping="Wrap"/>
      </StackPanel>

      <Grid Grid.Row="1" VerticalAlignment="Top">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="{left_col_w}"/>
          <ColumnDefinition Width="12"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="{bg_panel}" BorderBrush="{border}"
                BorderThickness="1" CornerRadius="4" Padding="{pad_panel}"
                VerticalAlignment="Top" HorizontalAlignment="Stretch">
          <StackPanel>
            <TextBlock Text="Esquema de la barra" Foreground="{fg_body}"
                       FontSize="11" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <TextBlock Text="Refleja cortes · clic en T# enfoca el editor · clic cerca de C# lo quita"
                       Foreground="{fg_muted}" FontSize="10" TextWrapping="Wrap"
                       Margin="0,0,0,6"/>
            <Border Background="{bg_input}" BorderBrush="{border}" BorderThickness="1"
                    CornerRadius="4" Padding="4" HorizontalAlignment="Center"
                    VerticalAlignment="Top">
              <Canvas x:Name="CnvBar" Width="{canvas_w}" Height="{canvas_h}"
                      Background="{bg_input}" Cursor="Arrow"
                      ClipToBounds="False"/>
            </Border>
          </StackPanel>
        </Border>

        <Border Grid.Column="2" Background="{bg_panel}" BorderBrush="{border}"
                BorderThickness="1" CornerRadius="4" Padding="0"
                VerticalAlignment="Top" HorizontalAlignment="Stretch"
                MinHeight="{schema_block_h}">
          <ScrollViewer VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Disabled"
                        MaxHeight="{schema_block_h}">
            <StackPanel Margin="{pad_panel}">
              <TextBlock Style="{{StaticResource LabelSmall}}"
                         Text="Grado de hormigón (tabla traslape)" Margin="0,0,0,4"/>
              <Grid Margin="0,0,0,10">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <ComboBox x:Name="CmbGrade" Grid.Column="0"
                          Style="{{StaticResource Combo}}" Height="28"/>
                <Button x:Name="BtnPickCuts" Grid.Column="1" Content="Marcar en vista"
                        Style="{{StaticResource BtnSelectOutline}}" MinWidth="120"
                        Margin="8,0,0,0"
                        ToolTip="Oculta la UI; clic sobre la barra (marca temporal). Finalizar en la barra de opciones"/>
                <Button x:Name="BtnLimpiar" Grid.Column="2" Content="Limpiar cortes"
                        Style="{{StaticResource BtnSelectOutline}}" MinWidth="110"
                        Margin="8,0,0,0"/>
              </Grid>

              <Border Background="{bg_input}" BorderBrush="{border}" BorderThickness="1"
                      CornerRadius="4" Padding="10" Margin="0,0,0,10">
                <TextBlock x:Name="TxtMeta" TextWrapping="Wrap" Foreground="{fg_body}"
                           FontSize="11" Text=""/>
              </Border>

              <TextBlock x:Name="TxtAlert" TextWrapping="Wrap" FontSize="11"
                         Foreground="{fg_body}" Margin="0,0,0,10"/>

              <TextBlock Text="Modo de traslape" Style="{{StaticResource LabelSmall}}"
                         Margin="0,0,0,4"/>
              <ComboBox x:Name="CmbLapMode" Style="{{StaticResource Combo}}"
                        Height="28" Margin="0,0,0,4"/>
              <TextBlock x:Name="TxtLapModeHint" TextWrapping="Wrap"
                         Foreground="{fg_muted}" FontSize="10" Margin="0,0,0,10"
                         Text=""/>

              <TextBlock Text="Largos de tramo (mm)" Style="{{StaticResource LabelSmall}}"
                         Margin="0,0,0,4"/>
              <Border x:Name="BrdTramosHost" Background="Transparent"
                      MinHeight="96" Margin="0,0,0,4"/>
            </StackPanel>
          </ScrollViewer>
        </Border>
      </Grid>

      <Grid Grid.Row="2" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="TxtStatus" Grid.Column="0" VerticalAlignment="Center"
                   Foreground="{fg_muted}" FontSize="{font_status}"
                   TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnApply" Content="Aplicar"
                  Style="{{StaticResource BtnPrimary}}" MinWidth="120" Margin="0,0,8,0"/>
          <Button x:Name="BtnClose" Content="Cerrar"
                  Style="{{StaticResource BtnSelectOutline}}" MinWidth="90"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
""".format(
        chrome_title=WINDOW_CHROME_TITLE,
        bg_app=BG_APP,
        border=BORDER,
        pad_window=PAD_WINDOW,
        font_family=FONT_FAMILY,
        font_base=FONT_SIZE_BASE,
        fg_title=FG_TITLE,
        font_title=FONT_SIZE_TITLE,
        font_weight_title=FONT_WEIGHT_TITLE,
        fg_body=FG_BODY,
        font_subtitle=FONT_SIZE_SUBTITLE,
        fg_muted=FG_MUTED,
        font_status=FONT_SIZE_STATUS,
        bg_panel=BG_PANEL,
        pad_panel=PAD_PANEL,
        bg_input=BG_INPUT,
        title_esc=title_esc,
        styles=BIMTOOLS_DARK_STYLES_XML,
        canvas_w=int(_CANVAS_W),
        canvas_h=int(_CANVAS_H),
        left_col_w=int(_LEFT_COL_W),
        schema_block_h=int(_SCHEMA_BLOCK_H),
        window_w=int(_WINDOW_W),
        window_h=int(_WINDOW_H),
        window_min_w=int(_WINDOW_MIN_W),
        window_min_h=int(_WINDOW_MIN_H),
    )
    try:
        AppDomain.CurrentDomain.SetData(_XAML_AD_KEY, _XAML_CACHE)
    except Exception:
        pass
    return _XAML_CACHE


def _prepare_window(win, uiapp):
    if win is None:
        return
    # Solo centrar una vez (sin bind Loaded/ContentRendered que reubica y
    # puede pelear el z-order al pintar el canvas).
    # No Owner→Revit: si Revit «No responde», Windows oculta las ventanas hijas.
    try:
        hwnd = revit_main_hwnd(uiapp)
        position_wpf_window_center_on_monitor(win, hwnd)
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
        if win.WindowState == WindowState.Minimized:
            win.WindowState = WindowState.Normal
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
        if win.WindowState == WindowState.Minimized:
            win.WindowState = WindowState.Normal
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
        win = self._window_ref()
        if win is None:
            return
        try:
            win.apply_in_revit(uiapp)
        except Exception as ex:
            win._ui_set_status(_as_unicode(ex), err=True)
            win._set_busy(False)


class _PickCutsHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def GetName(self):
        return u"DividirRebarPunto.PickCuts"

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        try:
            win.pick_cuts_in_revit(uiapp)
        except Exception as ex:
            win._ui_set_status(_as_unicode(ex), err=True)
            win._set_busy(False)


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
        self._busy = False
        self._rebuilding = False
        self._schema_built = False
        self._plan_uv = list(self._session.get(u"plan_points_uv") or [])
        self._plan_arc_mm = list(self._session.get(u"plan_arc_mm") or [])
        self._plan_flip_v = bool(self._session.get(u"plan_flip_v", True))
        self._plan_px = []
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
        self._brd_tramos_host = self._win.FindName(u"BrdTramosHost")
        self._pnl_tramos = None
        self._ensure_tramos_panel()
        self._btn_limpiar = self._win.FindName(u"BtnLimpiar")
        self._btn_pick_cuts = self._win.FindName(u"BtnPickCuts")
        self._btn_apply = self._win.FindName(u"BtnApply")
        self._btn_close = self._win.FindName(u"BtnClose")

        if self._txt_title is not None:
            self._txt_title.Text = _TITULO
        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = (
                u"Marque cortes en la vista; luego afine el largo de cada tramo "
                u"y pulse Aplicar."
            )

        self._handler_apply = _ApplyHandler(weakref.ref(self))
        self._ext_apply = ExternalEvent.Create(self._handler_apply)
        self._handler_pick = _PickCutsHandler(weakref.ref(self))
        self._ext_pick = ExternalEvent.Create(self._handler_pick)

        self._populate_grade_combo()
        self._populate_lap_mode_combo()
        self._wire()
        self._update_meta()
        self._update_lap_mode_hint()

        # Esquema / paneles se montan en show() antes de Show (sin shell vacío).
        self._ui_set_alert(
            u"Pulse «Marcar en vista»: la UI se oculta; marque puntos sobre la barra "
            u"(marca temporal). Finalizar en la barra de opciones (arriba o abajo).",
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
        if self._btn_pick_cuts is not None:
            self._btn_pick_cuts.Click += RoutedEventHandler(self._on_pick_cuts)
        self._cnv.MouseLeftButtonDown += MouseButtonEventHandler(self._on_canvas_click)
        self._win.Closed += EventHandler(self._on_closed)
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
                self._btn_pick_cuts,
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
            if _as_unicode(plane) == u"normal":
                lines.append(u"Esquema: proyección en el plano de la barra")
            else:
                lines.append(
                    u"Esquema proyectado: plano {0}".format(
                        _as_unicode(plane).upper()
                    )
                )
        n_pos = self._session.get(u"n_positions") or 1
        layout_name = self._session.get(u"layout") or u""
        if n_pos > 1 or layout_name == u"MaximumSpacing":
            lines.append(layout_label(layout_name, n_pos))
        if self._txt_meta is not None:
            self._txt_meta.Text = u"\n".join(lines)

    def _rebuild_plan_px(self):
        self._plan_px, self._scale = fit_polyline_to_canvas(
            self._plan_uv,
            _CANVAS_W,
            _CANVAS_H,
            _MARGIN,
            swap_uv=False,
            flip_v=self._plan_flip_v,
        )
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
            return _MARGIN, _MARGIN
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
                lx = max(4.0, min(float(_CANVAS_W) - 44.0, lx))
                ly = max(4.0, min(float(_CANVAS_H) - 24.0, ly))
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
        try:
            intervals = piece_intervals_with_lap(
                self._total_mm,
                self._cuts_mm,
                self._lap_mm,
                lap_mode=self._selected_lap_mode(),
            )
            return [max(0.0, float(b) - float(a)) for a, b in intervals]
        except Exception:
            return []

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

        tb = TextBox()
        tb.Text = u"{0:.0f}".format(float(span.get(u"length_mm") or 0))
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

        if fab_len is not None:
            sub = TextBlock()
            sub.Text = u"Fabricado c/ traslape ≈ {0:.0f} mm".format(fab_len)
            sub.Foreground = _brush(FG_MUTED)
            sub.FontSize = 9.0
            block.Children.Add(sub)

        return block

    def _rebuild_tramos_panel(self):
        """Editor de largo por tramo (vano en centerline) + preview con traslape."""
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
                    u"Aún no hay cortes válidos. Pulse «Marcar en vista», "
                    u"marque puntos sobre la barra y Finalizar. "
                    u"Aquí aparecerán T1, T2… para afinar cada largo."
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
        Monta canvas + paneles con la ventana aún oculta (antes de Show).

        Usa tamaños fijos del XAML (_CANVAS_W/H); no depende de ActualWidth
        ni del evento Loaded. Guard `_schema_built` evita reentrada.
        """
        if self._schema_built:
            return
        self._schema_built = True
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
        ok, msg, new_cuts = set_span_length_mm(
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
            u"Tramo T{0} → {1:.0f} mm".format(span_index + 1, new_len), ok=True
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
        if self._busy:
            return
        pos = args.GetPosition(self._cnv)
        hit = self._nearest_cut_index_at_px(pos.X, pos.Y)
        if hit >= 0:
            self._remove_cut_at_index(hit)
            self._ui_set_status(u"Corte eliminado.")
            return
        span_i = self._span_index_at_px(pos.X, pos.Y)
        if span_i >= 0 and self._cuts_mm:
            self._set_active_tramo(span_i, focus_editor=True)
            self._ui_set_status(
                u"{0} seleccionado en esquema · afine el largo a la derecha.".format(
                    tramo_label(span_i)
                ),
                ok=True,
            )
            return
        self._ui_set_status(
            u"Use «Marcar en vista» para añadir cortes sobre la barra en Revit.",
            err=True,
        )

    def _on_limpiar(self, sender, args):
        self._cuts_mm = []
        self._refresh_all()
        self._ui_set_status(u"Cortes limpiados.")
        self._ui_set_alert(
            u"Pulse «Marcar en vista», marque puntos y Finalizar en la barra de opciones.",
            u"info",
        )

    def _hide_for_pick(self):
        try:
            self._win.Hide()
        except Exception:
            pass

    def _show_after_pick(self):
        def _do():
            try:
                if not bool(getattr(self._win, u"IsVisible", False)):
                    self._win.Show()
            except Exception:
                try:
                    self._win.Show()
                except Exception:
                    pass
            try:
                if self._win.WindowState == WindowState.Minimized:
                    self._win.WindowState = WindowState.Normal
            except Exception:
                pass
            try:
                self._win.Activate()
                self._win.Focus()
            except Exception:
                pass

        self._invoke_ui_sync(_do)

    def _on_pick_cuts(self, sender, args):
        if self._busy:
            return
        if self._rebar_id is None:
            self._ui_set_status(u"No hay barra activa.", err=True)
            return
        self._set_busy(
            True,
            u"Marque cortes en la vista (marca temporal). Finalizar en la barra de opciones…",
        )
        self._hide_for_pick()
        self._ext_pick.Raise()

    def pick_cuts_in_revit(self, uiapp):
        uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
        if uidoc is None:
            self._ui_set_status(u"No hay documento activo.", err=True)
            self._show_after_pick()
            self._set_busy(False)
            return
        doc = uidoc.Document
        rebar = doc.GetElement(self._rebar_id) if self._rebar_id is not None else None
        if not isinstance(rebar, Rebar):
            self._ui_set_status(u"La barra ya no está disponible.", err=True)
            self._show_after_pick()
            self._set_busy(False)
            return

        added = 0
        skipped = 0
        cancelled = False
        last_reject = u""
        try:
            points, err = pick_cut_points_on_rebar(
                uidoc,
                rebar,
                prompt=(
                    u"Clic sobre la barra en cada punto de corte. "
                    u"Luego Finalizar en la barra de opciones (arriba o abajo)."
                ),
                lap_mm=self._lap_mm,
                existing_cuts_mm=list(self._cuts_mm),
            )
            if points is None:
                cancelled = True
                msg = _as_unicode(err or u"")
                if u"cancel" not in msg.lower() and u"Selección cancelada" not in msg:
                    self._ui_set_status(msg or u"Selección cancelada.", err=True)
            else:
                for pt in points:
                    ok_p, msg_p, cut_mm = cut_mm_from_pick(doc, rebar, pt)
                    if not ok_p or cut_mm is None:
                        skipped += 1
                        if msg_p:
                            last_reject = _as_unicode(msg_p)
                        continue
                    mm_val = float(cut_mm)
                    result = [False]
                    reject_msg = [u""]

                    def _add(mm=mm_val, res=result, rmsg=reject_msg):
                        trial = sorted(self._cuts_mm + [float(mm)])
                        ok_c, msg_c, sorted_cuts = validate_cuts_with_lap(
                            self._total_mm,
                            trial,
                            self._lap_mm,
                            _MIN_PIECE_MM,
                            lap_mode=self._selected_lap_mode(),
                        )
                        if not ok_c:
                            rmsg[0] = _as_unicode(msg_c or u"")
                            res[0] = False
                            return
                        self._cuts_mm = list(sorted_cuts)
                        self._refresh_all(update_status=False)
                        res[0] = True

                    self._invoke_ui_sync(_add)
                    if result[0]:
                        added += 1
                    else:
                        skipped += 1
                        if reject_msg[0]:
                            last_reject = reject_msg[0]
        finally:
            try:
                _clear_selection(uidoc)
            except Exception:
                pass
            self._show_after_pick()
            # Asegura editores de tramo visibles tras el pick (cortes previos o nuevos).
            try:
                self._invoke_ui_sync(
                    lambda: self._rebuild_tramos_panel()
                )
            except Exception:
                pass
            if cancelled and added == 0:
                self._ui_set_status(u"Selección cancelada.")
                self._ui_set_alert(
                    u"Pulse «Marcar en vista» para volver a marcar cortes.",
                    u"info",
                )
            elif added > 0:
                extra = u""
                if skipped > 0:
                    extra = u" ({0} punto(s) omitidos por validación)".format(skipped)
                self._ui_set_alert(
                    u"{0} corte(s) marcados en vista{1}. Afine largos de tramo si hace falta y pulse Aplicar.".format(
                        added, extra
                    ),
                    u"ok",
                )
                self._ui_set_status(
                    u"Marcados {0} corte(s) · total {1}".format(
                        added, len(self._cuts_mm)
                    ),
                    ok=True,
                )
            else:
                if skipped > 0:
                    detail = last_reject or u"revise distancia a extremos (≥ lap/2) y entre cortes (≥ lap)"
                    self._ui_set_status(
                        u"Ningún corte válido ({0} punto(s) rechazados). {1}".format(
                            skipped, detail
                        ),
                        err=True,
                    )
                    if self._cuts_mm:
                        self._ui_set_alert(
                            u"Se mantienen {0} corte(s) previos. Afine largos abajo o marque de nuevo.".format(
                                len(self._cuts_mm)
                            ),
                            u"warn",
                        )
                else:
                    self._ui_set_status(
                        u"Sin puntos: marque sobre la barra y pulse Finalizar.",
                        err=True,
                    )
            self._set_busy(False)

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
        if self._busy:
            return
        if not self._cuts_mm:
            self._ui_set_status(u"Añada al menos un corte.", err=True)
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
        self._set_busy(True, u"Aplicando división en Revit…")
        self._ext_apply.Raise()

    def apply_in_revit(self, uiapp):
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            self._ui_set_status(u"No hay documento activo.", err=True)
            self._set_busy(False)
            return
        doc = uidoc.Document
        rebar = doc.GetElement(self._rebar_id)
        if not isinstance(rebar, Rebar):
            self._ui_set_status(u"La barra ya no existe en el modelo.", err=True)
            self._set_busy(False)
            return
        target_view = resolve_active_model_view(uidoc)
        view_for_unobscured = target_view
        if target_view is not None:
            try:
                view_for_unobscured = doc.GetElement(target_view.Id)
            except Exception:
                view_for_unobscured = resolve_active_model_view(uidoc)

        ok, msg, ids, meta = divide_rebar_at_cuts(
            doc,
            rebar,
            sorted(self._cuts_mm),
            self._selected_grade(),
            view=view_for_unobscured,
            lap_mode=self._selected_lap_mode(),
        )
        if ok:
            try:
                _clear_selection(uidoc)
            except Exception:
                pass
            try:
                self._win.Close()
            except Exception:
                _clear_singleton()
            return

        self._set_busy(False)
        self._ui_set_alert(msg, u"warn")
        self._ui_set_status(msg, err=True)
        mostrar_aviso(uiapp, u"No se pudo dividir.", msg)

    def _on_close(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass

    def _on_closed(self, sender, args):
        _clear_singleton()

    def show(self):
        """Centrar + montar esquema oculto → Show una sola vez con UI completa."""
        _prepare_window(self._win, self._uiapp)
        self._build_schema_ready()
        self._win.Show()
        _bring_window_to_front(self._win)


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
        uidoc.Document, rebar, None
    )
    if not ok:
        mostrar_aviso(revit, prep_err or u"Barra no elegible.")
        return
    win = DividirRebarPuntoWindow(revit, rebar, session)
    _set_singleton(win)
    win.show()
