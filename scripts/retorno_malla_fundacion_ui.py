# -*- coding: utf-8 -*-
"""
UI WPF — Arainco: Remate Mallas.

Elevación + sección + rail configuradores (mockup validado).
Hasta 3 capas · extremos Pata L / Empotramiento · host WallFoundation.
"""

from __future__ import print_function

import weakref

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import (
    RoutedEventHandler,
    SizeChangedEventHandler,
    Thickness,
    VerticalAlignment,
    Visibility,
    WindowStartupLocation,
)
from System.Windows.Controls import (
    Border,
    Button,
    Canvas,
    ComboBox,
    ComboBoxItem,
    Orientation,
    StackPanel,
    TextBlock,
)
from System.Windows.Input import MouseButtonEventHandler
from System.Windows.Media import Color, SolidColorBrush
from System.Windows.Shapes import Ellipse, Line, Rectangle
from System.Windows.Markup import XamlReader
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

from bimtools_instruction_dialog import show_message_dialog
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from barras_retorno_malla_geom import (
    CAPAS_OPTS,
    COVER_FUND_BOT_MM,
    COVER_WALL_BOT_MM,
    COVER_WALL_MM,
    DIAMS_MM,
    DOSIFICACION_HORMIGON_DEFAULT,
    DOSIFICACION_HORMIGON_OPCIONES,
    END_EMPOTRO,
    END_PATA,
    LAP_MODE_LABELS,
    LAP_MODE_SYMMETRIC,
    LAYER_CENTERLINE_SPACING_MM,
    MARGIN_END_MM,
    MAX_BARRA_COMERCIAL_MM,
    N_BARS_OPTS,
    clamp_n_capas,
    cover_axis_offset_mm_for_layer,
    empotramiento_mm_from_diam,
    format_mm_es,
    lap_zone_around_cut,
    layer_bar_offsets_mm,
    normalize_concrete_grade,
    normalize_end_condition,
    normalize_lap_mode_ui,
    pata_mm_from_diam,
    stagger_cuts_for_layer,
    sync_layers,
    toggle_cut_at_mm,
    traslape_mm_from_diam,
)
from barras_retorno_malla_place import (
    wall_elev_canvas_flip_for_view,
    wall_meta_for_ui,
)
from retorno_malla_fundacion_geom import (
    COVER_SUPERIOR_MM,
    MODE_INFERIOR_FUND,
    MODE_OPTS,
    MODE_SUPERIOR,
    mode_label,
    normalize_mode,
)
from retorno_malla_fundacion_place import place_retorno_malla_fundacion
from revit_wpf_window_position import revit_main_hwnd

_DIALOG_TITLE = u"Arainco: Remate Mallas"
_SINGLETON_KEY = u"Arainco.RetornoMallaHostFundacion.ActiveWindow"
_END_OPTS_UI = (
    (END_PATA, u"Pata L"),
    (END_EMPOTRO, u"Empotramiento"),
)
_HOST_STROKE = u"#c084fc"
_BAR = u"#fbbf24"
_CUT = u"#f87171"
_WALL = u"#2a4a58"
_WALL_STROKE = u"#5a8a9a"
_FUND = u"#1a3544"
_FUND_STROKE = u"#4a7a88"
_JOINED = u"#1e3a4a"
_JOINED_STROKE = u"#3d6a7a"
_ACCENT = u"#5BC0DE"
_LAYER_COLORS = (u"#fbbf24", u"#38bdf8", u"#a78bfa")
_TRAMO_COLORS = (u"#38bdf8", u"#a78bfa", u"#34d399", u"#fb923c")

_UI_WIDTH = 1080
_UI_HEIGHT = 980
_UI_SIDE_W = 380
_UI_SECTION_H = 132
_STROKE_BAR = 1.6
_STROKE_WALL = 1.5
_STROKE_CUT = 1.4


def _brush(hex_color, alpha=255):
    h = (hex_color or u"#95B8CC").lstrip(u"#")
    if len(h) != 6:
        return SolidColorBrush(Color.FromRgb(149, 184, 204))
    return SolidColorBrush(
        Color.FromArgb(alpha, int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    )


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _mostrar_aviso(uiapp, instruction, content=u""):
    try:
        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            _DIALOG_TITLE,
            instruction=_as_unicode(instruction),
            content=_as_unicode(content) if content else None,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        from Autodesk.Revit.UI import TaskDialog

        TaskDialog.Show(
            _DIALOG_TITLE,
            u"{0}\n{1}".format(_as_unicode(instruction), _as_unicode(content)).strip(),
        )
    except Exception:
        pass


def _clear_singleton():
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
    except Exception:
        pass


def _find_live_window_by_title():
    try:
        from System.Windows import Application

        target = _DIALOG_TITLE
        for win in Application.Current.Windows:
            try:
                if _as_unicode(win.Title) != target:
                    continue
                if hasattr(win, "IsLoaded") and not bool(win.IsLoaded):
                    continue
                return win
            except Exception:
                continue
    except Exception:
        pass
    return None


def _get_active_window():
    win = None
    try:
        win = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        win = None
    if win is not None:
        try:
            _ = win.Title
        except Exception:
            win = None
            _clear_singleton()
        if win is not None:
            try:
                if hasattr(win, "IsLoaded") and not bool(win.IsLoaded):
                    win = None
                    _clear_singleton()
            except Exception:
                pass
            if win is not None:
                try:
                    if hasattr(win, "IsVisible") and not bool(win.IsVisible):
                        win = None
                        _clear_singleton()
                except Exception:
                    pass
    if win is None:
        win = _find_live_window_by_title()
        if win is not None:
            try:
                AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
            except Exception:
                pass
    return win


def _focus_existing():
    win = _get_active_window()
    if win is None:
        return False
    try:
        if win.WindowState == getattr(win.WindowState, u"Minimized", win.WindowState):
            from System.Windows import WindowState

            win.WindowState = WindowState.Normal
    except Exception:
        pass
    try:
        win.Activate()
        win.Focus()
    except Exception:
        pass
    _mostrar_aviso(None, u"La herramienta ya esta en ejecucion.")
    return True


def _end_label_ui(key):
    k = normalize_end_condition(key)
    if k == END_PATA:
        return u"Pata L"
    return u"Empotramiento"


class _ColocarHandler(IExternalEventHandler):
    def __init__(self, win_ref):
        self._win_ref = win_ref

    def GetName(self):
        return u"AraincoRetornoMallaHostFundacionColocar"

    def Execute(self, app):
        win = self._win_ref() if self._win_ref else None
        if win is None:
            return
        try:
            win._execute_colocar()
        except Exception as ex:
            try:
                win._set_status(_as_unicode(ex))
            except Exception:
                pass


class RetornoMallaHostFundacionWindow(object):
    def __init__(self, uiapp, uidoc, doc, walls, on_closed=None):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        self._walls = list(walls or [])
        self._wall = self._walls[0] if self._walls else None
        self._on_closed = on_closed
        self._active_view = None
        try:
            if uidoc is not None:
                self._active_view = uidoc.ActiveView
        except Exception:
            self._active_view = None
        self._meta = (
            wall_meta_for_ui(doc, self._wall, view=self._active_view)
            if self._wall
            else {}
        )
        self._elev_flip = bool(
            (self._meta or {}).get(u"elev_flip")
            or wall_elev_canvas_flip_for_view(self._wall, self._active_view)
        )
        self._layers = sync_layers(None, 1)
        self._n_capas = 1
        self._lap_mode = LAP_MODE_SYMMETRIC
        self._concrete_grade = DOSIFICACION_HORMIGON_DEFAULT
        self._end_a = END_PATA
        self._end_b = END_PATA
        self._cuts_mm = []
        self._mode = MODE_INFERIOR_FUND
        self._cmb_capas = None
        self._cmb_modo = None
        self._cmb_lap = None
        self._cmb_dosif = None
        self._cmb_end_a = None
        self._cmb_end_b = None
        self._cmb_end_left = None
        self._cmb_end_right = None
        self._txt_end_left = None
        self._txt_end_right = None
        self._elev_flip = bool(getattr(self, u"_elev_flip", False))
        self._layer_host = None
        self._elev_draw = {}
        self._canvas = None
        self._canvas_sec = None
        self._txt_status = None
        self._txt_warn = None
        self._warn_border = None
        self._no_fund_border = None
        self._btn_colocar = None
        self._layer_combos = []
        self._closed = False

        self._colocar_handler = _ColocarHandler(weakref.ref(self))
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)

        body = self._build_body_xaml()
        footer = (
            u'<Button x:Name="BtnCancelar" Content="Cancelar" '
            u'Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>'
            u'<Button x:Name="BtnColocar" Content="Colocar armadura" '
            u'Style="{StaticResource BtnPrimary}" MinWidth="160"/>'
        )
        hint = u"Configure capas y extremos · clic en elevación para empalmes"
        xaml = build_simple_tool_xaml(
            title=_DIALOG_TITLE,
            styles_xml=BIMTOOLS_DARK_STYLES_XML,
            body_xaml=body,
            footer_actions_xaml=footer,
            footer_hint_xaml=hint,
            width=_UI_WIDTH,
            min_width=_UI_WIDTH,
            height=_UI_HEIGHT,
            min_height=_UI_HEIGHT,
            resize_mode=u"NoResize",
            size_to_content_height=False,
        )
        self._win = XamlReader.Parse(xaml)
        try:
            hwnd = revit_main_hwnd(uiapp)
            if hwnd is not None:
                from System.Windows.Interop import WindowInteropHelper

                WindowInteropHelper(self._win).Owner = hwnd
        except Exception:
            pass
        self._win.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self._apply_fixed_window_size()
        self._wire()
        self._rebuild_layer_cards()
        self._redraw()
        self._redraw_section()
        self._refresh_warn()
        self._set_status(self._status_line())
        try:
            self._win.Loaded += RoutedEventHandler(self._on_loaded)
        except Exception:
            pass

        def _on_closed(sender, args):
            self._closed = True
            _clear_singleton()
            if self._on_closed:
                try:
                    self._on_closed()
                except Exception:
                    pass

        self._win.Closed += EventHandler(_on_closed)
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, self._win)
        except Exception:
            pass

    def _has_fund(self):
        return bool(self._meta.get(u"foundation"))

    def _is_superior_mode(self):
        return normalize_mode(self._mode) == MODE_SUPERIOR

    def _read_mode(self):
        cmb = self._cmb_modo
        if cmb is None:
            return normalize_mode(self._mode)
        try:
            si = cmb.SelectedItem
            if si is not None:
                return normalize_mode(getattr(si, u"Tag", None) or si.Content)
        except Exception:
            pass
        return normalize_mode(self._mode)

    def _main_mm(self):
        try:
            return float(self._meta.get(u"main_mm") or 0.0)
        except Exception:
            return 0.0

    def _build_body_xaml(self):
        xaml = u"""
<Grid>
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="*"/>
    <ColumnDefinition Width="__SIDE_W__"/>
  </Grid.ColumnDefinitions>

  <!-- Elevación (preview principal) -->
  <Border Grid.Column="0" Background="#050E18" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" Margin="0,0,10,0">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <DockPanel Grid.Row="0" Margin="10,8,10,2">
        <Border DockPanel.Dock="Right" BorderBrush="#5BC0DE" BorderThickness="1"
                Background="#0E1B32" Padding="7,2" CornerRadius="3"
                ToolTip="Armadura longitudinal">
          <TextBlock Text="LONG." Foreground="#5BC0DE" FontSize="10" FontWeight="SemiBold"/>
        </Border>
        <TextBlock x:Name="TxtElevHdr"
                   Text="ELEVACIÓN · MURO"
                   Foreground="#E8F4F8" FontSize="12" FontWeight="SemiBold"
                   VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
      </DockPanel>
      <Border Grid.Row="1" Background="#0E1B32" BorderBrush="#21465C" BorderThickness="0,0,0,1"
              Padding="10,5" Margin="0,2,0,0">
        <TextBlock x:Name="TxtPickHint" Foreground="#fbbf24" FontSize="11"
                   TextWrapping="Wrap"
                   Text="Clic en la elevación para añadir o quitar empalmes (capa 1 de referencia; 2–3 alternan)."/>
      </Border>
      <Canvas x:Name="CnvElev" Grid.Row="2" Background="#050E18" ClipToBounds="True"
              Cursor="Cross"/>
    </Grid>
  </Border>

  <!-- Rail de configuración -->
  <Border Grid.Column="1" Background="#0a1620" BorderBrush="#21465C"
          BorderThickness="1" CornerRadius="4">
    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                  Padding="10,10,8,10" CanContentScroll="False">
      <StackPanel x:Name="PanelStack">

        <!-- Modo -->
        <TextBlock Text="MODO" Foreground="#64748b" FontSize="10" FontWeight="SemiBold"
                   Margin="0,0,0,4"/>
        <ComboBox x:Name="CmbModo" Style="{StaticResource ComboStretch}"
                  HorizontalAlignment="Stretch" Margin="0,0,0,6"/>
        <Border x:Name="HostCard" Background="#0E1B32" BorderBrush="#c084fc"
                BorderThickness="1" CornerRadius="4" Padding="8,6" Margin="0,0,0,10">
          <StackPanel>
            <TextBlock x:Name="TxtHostTitle" Foreground="#c084fc" FontSize="10"
                       FontWeight="SemiBold" Text="HOST · WallFoundation"/>
            <TextBlock x:Name="TxtHostCard" Foreground="#95B8CC" FontSize="10"
                       TextWrapping="Wrap" Margin="0,2,0,0"
                       Text="Geometría del muro · Z fondo fund. −50 mm"/>
          </StackPanel>
        </Border>

        <!-- Sección -->
        <TextBlock Text="SECCIÓN" Foreground="#64748b" FontSize="10" FontWeight="SemiBold"
                   Margin="0,0,0,4"/>
        <Border BorderBrush="#21465C" BorderThickness="1" CornerRadius="4" Margin="0,0,0,12">
          <Canvas x:Name="CnvSec" Background="#050E18" Height="__SEC_H__" ClipToBounds="True"/>
        </Border>

        <!-- Capas (bloque principal) -->
        <DockPanel Margin="0,0,0,6">
          <StackPanel DockPanel.Dock="Right" Orientation="Horizontal"
                      VerticalAlignment="Center">
            <TextBlock Text="Hormigón" Foreground="#95B8CC" FontSize="10"
                       VerticalAlignment="Center" Margin="0,0,6,0"/>
            <ComboBox x:Name="CmbDosificacionHormigon" Style="{StaticResource Combo}"
                      MinWidth="68" VerticalAlignment="Center"
                      IsEditable="False" IsReadOnly="True"
                      ToolTip="Dosificación del hormigón (traslape)"/>
          </StackPanel>
          <TextBlock Text="CAPAS" Foreground="#E8F4F8" FontSize="12"
                     FontWeight="SemiBold" VerticalAlignment="Center"/>
        </DockPanel>
        <DockPanel Margin="0,0,0,8">
          <TextBlock DockPanel.Dock="Left" Text="Cantidad (máx. 3)"
                     Foreground="#95B8CC" FontSize="11" VerticalAlignment="Center"
                     Margin="0,0,10,0"/>
          <ComboBox x:Name="CmbCapas" Style="{StaticResource ComboStretch}"
                    HorizontalAlignment="Stretch"/>
        </DockPanel>
        <StackPanel x:Name="PanelLayers"/>

        <Border Height="1" Background="#21465C" Margin="0,4,0,10"/>

        <!-- Extremos -->
        <TextBlock Text="EXTREMOS" Foreground="#E8F4F8" FontSize="12"
                   FontWeight="SemiBold" Margin="0,0,0,6"/>
        <Grid Margin="0,0,0,10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="8"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0">
            <TextBlock x:Name="TxtEndLeft" Text="Izq. vista (A)" Foreground="#95B8CC" FontSize="10" Margin="0,0,0,3"/>
            <ComboBox x:Name="CmbEndLeft" Style="{StaticResource ComboStretch}"
                      HorizontalAlignment="Stretch"/>
          </StackPanel>
          <StackPanel Grid.Column="2">
            <TextBlock x:Name="TxtEndRight" Text="Der. vista (B)" Foreground="#95B8CC" FontSize="10" Margin="0,0,0,3"/>
            <ComboBox x:Name="CmbEndRight" Style="{StaticResource ComboStretch}"
                      HorizontalAlignment="Stretch"/>
          </StackPanel>
        </Grid>

        <Border Height="1" Background="#21465C" Margin="0,0,0,10"/>

        <!-- Empalme -->
        <TextBlock Text="EMPALME" Foreground="#E8F4F8" FontSize="12"
                   FontWeight="SemiBold" Margin="0,0,0,4"/>
        <TextBlock Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,0,6"
                   Text="Tipo de traslape al dividir. Los cortes se marcan en la elevación."/>
        <ComboBox x:Name="CmbLap" Style="{StaticResource ComboStretch}"
                  HorizontalAlignment="Stretch" Margin="0,0,0,8"/>

        <Border x:Name="NoFundBorder" Background="#1a0a0a" BorderBrush="#ef4444"
                BorderThickness="1" CornerRadius="4" Padding="8,6" Margin="0,4,0,0"
                Visibility="Collapsed">
          <TextBlock x:Name="TxtNoFund" Foreground="#fca5a5" FontSize="10" TextWrapping="Wrap"
                     Text="Este muro no tiene fundación corrida unida. Cambie a modo superior o elija otro muro."/>
        </Border>
        <Border x:Name="WarnBorder" Background="#1a1408" BorderBrush="#d97706"
                BorderThickness="1" CornerRadius="4" Padding="8,6" Margin="0,6,0,0"
                Visibility="Collapsed">
          <TextBlock x:Name="TxtWarn" Foreground="#fcd34d" FontSize="10" TextWrapping="Wrap"/>
        </Border>
      </StackPanel>
    </ScrollViewer>
  </Border>
</Grid>
"""
        xaml = xaml.replace(u"__SIDE_W__", u"{0}".format(int(_UI_SIDE_W)))
        xaml = xaml.replace(u"__SEC_H__", u"{0}".format(int(_UI_SECTION_H)))
        return xaml.strip()

    def _apply_fixed_window_size(self):
        """Tamaño fijo: no redimensionable; Width/Height/Min/Max coinciden."""
        win = self._win
        if win is None:
            return
        try:
            from System.Windows import ResizeMode, SizeToContent

            win.SizeToContent = SizeToContent.Manual
            win.ResizeMode = ResizeMode.NoResize
        except Exception:
            pass
        try:
            w = float(_UI_WIDTH)
            h = float(_UI_HEIGHT)
            win.Width = w
            win.Height = h
            win.MinWidth = w
            win.MaxWidth = w
            win.MinHeight = h
            win.MaxHeight = h
        except Exception:
            pass

    def _on_loaded(self, sender, args):
        try:
            self._apply_fixed_window_size()
            self._redraw()
            self._redraw_section()
        except Exception:
            pass

    def _wire(self):
        w = self._win
        self._canvas = w.FindName(u"CnvElev")
        self._canvas_sec = w.FindName(u"CnvSec")
        self._cmb_capas = w.FindName(u"CmbCapas")
        self._cmb_modo = w.FindName(u"CmbModo")
        self._cmb_lap = w.FindName(u"CmbLap")
        self._cmb_dosif = w.FindName(u"CmbDosificacionHormigon")
        self._cmb_end_left = w.FindName(u"CmbEndLeft")
        self._cmb_end_right = w.FindName(u"CmbEndRight")
        self._txt_end_left = w.FindName(u"TxtEndLeft")
        self._txt_end_right = w.FindName(u"TxtEndRight")
        self._wire_end_combos_to_ab()
        self._layer_host = w.FindName(u"PanelLayers")
        self._txt_status = w.FindName(u"TxtStatus")
        self._txt_warn = w.FindName(u"TxtWarn")
        self._warn_border = w.FindName(u"WarnBorder")
        self._no_fund_border = w.FindName(u"NoFundBorder")
        self._btn_colocar = w.FindName(u"BtnColocar")

        self._update_elev_header()
        self._update_params_hint()
        self._update_host_card()

        if self._cmb_modo is not None:
            for key, lab in MODE_OPTS:
                it = ComboBoxItem()
                it.Content = lab
                it.Tag = key
                self._cmb_modo.Items.Add(it)
                if key == self._mode:
                    self._cmb_modo.SelectedItem = it
            if self._cmb_modo.SelectedItem is None and self._cmb_modo.Items.Count:
                self._cmb_modo.SelectedIndex = 0

        for n in CAPAS_OPTS:
            it = ComboBoxItem()
            it.Content = u"{0}".format(n)
            it.Tag = int(n)
            self._cmb_capas.Items.Add(it)
        self._cmb_capas.SelectedIndex = 0

        if self._cmb_dosif is not None:
            for lab in DOSIFICACION_HORMIGON_OPCIONES:
                it = ComboBoxItem()
                it.Content = lab
                it.Tag = lab
                self._cmb_dosif.Items.Add(it)
            self._cmb_dosif.SelectedIndex = 0

        def _fill_end_cmb(cmb, selected_key):
            if cmb is None:
                return
            for key, lab in _END_OPTS_UI:
                it = ComboBoxItem()
                it.Content = lab
                it.Tag = key
                cmb.Items.Add(it)
                if key == selected_key:
                    cmb.SelectedItem = it
            if cmb.SelectedItem is None and cmb.Items.Count:
                cmb.SelectedIndex = 0

        _fill_end_cmb(self._cmb_end_a, END_PATA)
        _fill_end_cmb(self._cmb_end_b, END_PATA)

        if self._cmb_lap is not None:
            for key, label in LAP_MODE_LABELS:
                it = ComboBoxItem()
                it.Content = label
                it.Tag = key
                self._cmb_lap.Items.Add(it)
            self._cmb_lap.SelectedIndex = 0

        from System.Windows.Controls import SelectionChangedEventHandler

        def on_capas(sender, args):
            it = self._cmb_capas.SelectedItem
            if it is None:
                return
            n = clamp_n_capas(it.Tag)
            self._n_capas = n
            self._layers = sync_layers(self._layers, n)
            self._rebuild_layer_cards()
            self._redraw()
            self._redraw_section()
            self._set_status(self._status_line())

        def on_lap(sender, args):
            it = self._cmb_lap.SelectedItem
            if it is None:
                return
            self._lap_mode = normalize_lap_mode_ui(it.Tag)
            self._redraw()
            self._set_status(self._status_line())

        def on_dosif(sender, args):
            self._concrete_grade = self._read_concrete_grade()
            self._sync_layer_tips()
            self._refresh_warn()
            self._redraw()
            self._set_status(self._status_line())

        def on_end(sender, args):
            self._end_a = self._read_end(self._cmb_end_a, END_PATA)
            self._end_b = self._read_end(self._cmb_end_b, END_PATA)
            self._redraw()
            self._set_status(self._status_line())

        self._cmb_capas.SelectionChanged += SelectionChangedEventHandler(on_capas)
        if self._cmb_modo is not None:
            def on_modo(sender, args):
                self._mode = self._read_mode()
                self._cuts_mm = []
                self._update_elev_header()
                self._update_params_hint()
                self._update_host_card()
                self._redraw()
                self._redraw_section()
                self._refresh_warn()
                self._set_status(self._status_line())

            self._cmb_modo.SelectionChanged += SelectionChangedEventHandler(on_modo)
        if self._cmb_dosif is not None:
            self._cmb_dosif.SelectionChanged += SelectionChangedEventHandler(on_dosif)
        if self._cmb_lap is not None:
            self._cmb_lap.SelectionChanged += SelectionChangedEventHandler(on_lap)
        if self._cmb_end_a is not None:
            self._cmb_end_a.SelectionChanged += SelectionChangedEventHandler(on_end)
        if self._cmb_end_b is not None:
            self._cmb_end_b.SelectionChanged += SelectionChangedEventHandler(on_end)

        if self._canvas is not None:
            def _on_size(sender, args):
                self._redraw()

            self._canvas.SizeChanged += SizeChangedEventHandler(_on_size)
            self._canvas.MouseLeftButtonDown += MouseButtonEventHandler(
                self._on_canvas_click
            )
        if self._canvas_sec is not None:
            def _on_sec_size(sender, args):
                self._redraw_section()

            self._canvas_sec.SizeChanged += SizeChangedEventHandler(_on_sec_size)

        btn_c = w.FindName(u"BtnCancelar")
        btn_p = w.FindName(u"BtnColocar")
        if btn_c is not None:
            btn_c.Click += RoutedEventHandler(lambda s, a: self._win.Close())
        if btn_p is not None:
            btn_p.Click += RoutedEventHandler(lambda s, a: self._colocar_event.Raise())

    def _read_concrete_grade(self):
        cmb = self._cmb_dosif
        if cmb is None:
            return DOSIFICACION_HORMIGON_DEFAULT
        try:
            si = cmb.SelectedItem
            if si is not None:
                tag = getattr(si, u"Tag", None)
                if tag is not None:
                    return normalize_concrete_grade(tag)
                return normalize_concrete_grade(si.Content)
        except Exception:
            pass
        return DOSIFICACION_HORMIGON_DEFAULT

    def _read_end(self, cmb, default):
        if cmb is None:
            return normalize_end_condition(default)
        try:
            si = cmb.SelectedItem
            if si is not None:
                return normalize_end_condition(getattr(si, u"Tag", None) or si.Content)
        except Exception:
            pass
        return normalize_end_condition(default)


    def _wire_end_combos_to_ab(self):
        """
        Combos izquierda/derecha = extremos en la vista activa.
        A/B = inicio/fin de LocationCurve (colocacion).
        Con elev_flip, izquierda vista = B y derecha = A.
        """
        flip = bool(getattr(self, u"_elev_flip", False))
        if flip:
            self._cmb_end_a = self._cmb_end_right
            self._cmb_end_b = self._cmb_end_left
            if self._txt_end_left is not None:
                self._txt_end_left.Text = u"Izq. vista (B)"
            if self._txt_end_right is not None:
                self._txt_end_right.Text = u"Der. vista (A)"
        else:
            self._cmb_end_a = self._cmb_end_left
            self._cmb_end_b = self._cmb_end_right
            if self._txt_end_left is not None:
                self._txt_end_left.Text = u"Izq. vista (A)"
            if self._txt_end_right is not None:
                self._txt_end_right.Text = u"Der. vista (B)"

    def _elev_flip_flag(self):
        return bool(
            getattr(self, u"_elev_flip", False)
            or (self._meta or {}).get(u"elev_flip")
        )

    def _update_elev_header(self):
        hdr = self._win.FindName(u"TxtElevHdr") if self._win else None
        if hdr is None:
            return
        orient = (
            u" · A a la derecha"
            if self._elev_flip_flag()
            else u" · A a la izquierda"
        )
        if self._is_superior_mode():
            hdr.Text = u"Elevación · tope muro (vista activa)" + orient
        elif self._has_fund():
            hdr.Text = u"Elevación · pie + fundación (vista activa)" + orient
        else:
            hdr.Text = u"Elevación · muro (vista activa)" + orient

    def _update_host_card(self):
        card = self._win.FindName(u"HostCard") if self._win else None
        title = self._win.FindName(u"TxtHostTitle") if self._win else None
        txt = self._win.FindName(u"TxtHostCard") if self._win else None
        if self._is_superior_mode():
            if card is not None:
                try:
                    card.BorderBrush = _brush(_ACCENT)
                except Exception:
                    pass
            if title is not None:
                title.Foreground = _brush(_ACCENT)
                title.Text = u"HOST · Muro"
            if txt is not None:
                txt.Text = (
                    u"Retorno en tope (coronamiento) · recub. superior {0} mm · "
                    u"distribución en espesor"
                ).format(int(COVER_SUPERIOR_MM))
        else:
            if card is not None:
                try:
                    card.BorderBrush = _brush(_HOST_STROKE)
                except Exception:
                    pass
            if title is not None:
                title.Foreground = _brush(_HOST_STROKE)
                title.Text = u"HOST · WallFoundation"
            if txt is not None:
                if self._has_fund():
                    txt.Text = (
                        u"Geometría del muro · Z fondo fund. −{0} mm · "
                        u"barras en espesor (n ≥ 2)"
                    ).format(int(COVER_FUND_BOT_MM))
                else:
                    txt.Text = (
                        u"Se requiere fundación corrida unida al muro para este modo."
                    )

    def _update_params_hint(self):
        # Host card concentra el contexto; se mantiene por compatibilidad de llamadas.
        return

    def _sync_layer_tips(self):
        grade = self._concrete_grade
        for i, triple in enumerate(self._layer_combos or []):
            tip = triple[2] if len(triple) > 2 else None
            if tip is None:
                continue
            d = 10
            try:
                if self._layers and i < len(self._layers):
                    d = self._layers[i].get(u"diam_mm", 10)
            except Exception:
                pass
            tip.Text = u"Traslape ≈ {0} mm ({1})".format(
                format_mm_es(traslape_mm_from_diam(d, grade)), grade
            )

    def _rebuild_layer_cards(self):
        host = self._layer_host
        if host is None:
            return
        host.Children.Clear()
        self._layer_combos = []
        from System.Windows.Controls import SelectionChangedEventHandler

        multi = len(self._layers) > 1
        for i, ly in enumerate(self._layers):
            color = _LAYER_COLORS[i % len(_LAYER_COLORS)]
            card = Border()
            card.Background = _brush(u"#050E18")
            card.BorderBrush = _brush(color, 160)
            card.BorderThickness = Thickness(1)
            try:
                from System.Windows import CornerRadius

                card.CornerRadius = CornerRadius(4)
            except Exception:
                pass
            card.Padding = Thickness(8, 7, 8, 7)
            card.Margin = Thickness(0, 0, 0, 8)

            sp = StackPanel()
            hdr_row = StackPanel()
            hdr_row.Orientation = Orientation.Horizontal

            accent = Border()
            accent.Width = 3
            accent.Height = 14
            accent.Background = _brush(color)
            accent.Margin = Thickness(0, 0, 6, 0)
            accent.VerticalAlignment = VerticalAlignment.Center
            try:
                from System.Windows import CornerRadius

                accent.CornerRadius = CornerRadius(1)
            except Exception:
                pass
            hdr_row.Children.Add(accent)

            hdr = TextBlock()
            if multi and i == 0:
                role = u"referencia (pick)"
            elif multi and i % 2 == 1:
                role = u"alterna A-B-A"
            elif multi:
                role = u"referencia"
            else:
                role = u"única"
            hdr.Text = u"Capa {0} · {1}".format(i + 1, role)
            hdr.Foreground = _brush(color)
            hdr.FontSize = 11
            try:
                from System.Windows import FontWeights

                hdr.FontWeight = FontWeights.SemiBold
            except Exception:
                pass
            hdr.VerticalAlignment = VerticalAlignment.Center
            hdr_row.Children.Add(hdr)
            sp.Children.Add(hdr_row)

            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Margin = Thickness(0, 6, 0, 0)

            def _mk_combo(values, selected, width=70, diam=False):
                cmb = ComboBox()
                try:
                    cmb.Style = self._win.FindResource(u"Combo")
                except Exception:
                    pass
                cmb.Width = width
                cmb.Margin = Thickness(0, 0, 8, 0)
                for v in values:
                    it = ComboBoxItem()
                    it.Content = u"Ø{0}".format(v) if diam else u"{0}".format(v)
                    it.Tag = int(v)
                    cmb.Items.Add(it)
                    if int(v) == int(selected):
                        cmb.SelectedItem = it
                if cmb.SelectedItem is None and cmb.Items.Count:
                    cmb.SelectedIndex = 0
                return cmb

            lbl_n = TextBlock()
            lbl_n.Text = u"Barras"
            lbl_n.Foreground = _brush(u"#95B8CC")
            lbl_n.FontSize = 10
            lbl_n.Margin = Thickness(0, 0, 4, 0)
            lbl_n.VerticalAlignment = VerticalAlignment.Center
            cmb_n = _mk_combo(N_BARS_OPTS, ly.get(u"n_bars", 2), 52)
            lbl_d = TextBlock()
            lbl_d.Text = u"Ø"
            lbl_d.Foreground = _brush(u"#95B8CC")
            lbl_d.FontSize = 10
            lbl_d.Margin = Thickness(4, 0, 4, 0)
            lbl_d.VerticalAlignment = VerticalAlignment.Center
            cmb_d = _mk_combo(DIAMS_MM, ly.get(u"diam_mm", 10), 72, diam=True)

            tip = TextBlock()
            tip.Foreground = _brush(u"#64748b")
            tip.FontSize = 10
            tip.Margin = Thickness(0, 5, 0, 0)

            li = i

            def _on_change(sender, args, idx=li):
                self._read_layer_combos()
                self._sync_layer_tips()
                self._redraw()
                self._redraw_section()
                self._refresh_warn()
                self._set_status(self._status_line())

            cmb_n.SelectionChanged += SelectionChangedEventHandler(_on_change)
            cmb_d.SelectionChanged += SelectionChangedEventHandler(_on_change)

            row.Children.Add(lbl_n)
            row.Children.Add(cmb_n)
            row.Children.Add(lbl_d)
            row.Children.Add(cmb_d)
            sp.Children.Add(row)
            sp.Children.Add(tip)
            card.Child = sp
            host.Children.Add(card)
            self._layer_combos.append((cmb_n, cmb_d, tip))
        self._sync_layer_tips()

    def _read_layer_combos(self):
        layers = []
        for cmb_n, cmb_d, _tip in self._layer_combos or []:
            n = 2
            d = 10
            try:
                if cmb_n.SelectedItem is not None:
                    n = int(cmb_n.SelectedItem.Tag)
            except Exception:
                pass
            try:
                if cmb_d.SelectedItem is not None:
                    d = int(cmb_d.SelectedItem.Tag)
            except Exception:
                pass
            layers.append({u"n_bars": n, u"diam_mm": d})
        if layers:
            self._layers = sync_layers(layers, len(layers))
            self._n_capas = len(self._layers)

    def _status_line(self):
        try:
            wid = self._meta.get(u"id") or u"?"
        except Exception:
            wid = u"?"
        parts = []
        for i, ly in enumerate(self._layers):
            parts.append(
                u"C{0} Ø{1}×{2}".format(
                    i + 1, ly.get(u"diam_mm", 10), ly.get(u"n_bars", 2)
                )
            )
        n_cuts = len(self._cuts_mm)
        cuts_txt = (
            u"sin cortes"
            if n_cuts <= 0
            else (u"1 corte" if n_cuts == 1 else u"{0} cortes".format(n_cuts))
        )
        return u"Muro {0} · {1} · {2} · {3} → {4} · {5}".format(
            wid,
            mode_label(self._mode),
            u" · ".join(parts) if parts else u"—",
            _end_label_ui(self._end_a),
            _end_label_ui(self._end_b),
            cuts_txt,
        )

    def _set_status(self, text):
        if self._txt_status is not None:
            try:
                self._txt_status.Text = _as_unicode(text)
            except Exception:
                pass

    def _refresh_no_fund(self):
        show = (not self._is_superior_mode()) and (not self._has_fund())
        if self._no_fund_border is not None:
            self._no_fund_border.Visibility = (
                Visibility.Visible if show else Visibility.Collapsed
            )
        if self._btn_colocar is not None:
            try:
                self._btn_colocar.IsEnabled = not show
            except Exception:
                pass

    def _refresh_warn(self):
        self._refresh_no_fund()
        main = self._main_mm()
        show = main > MAX_BARRA_COMERCIAL_MM
        if self._warn_border is not None:
            self._warn_border.Visibility = (
                Visibility.Visible if show else Visibility.Collapsed
            )
        if show and self._txt_warn is not None:
            self._txt_warn.Text = (
                u"Aviso: L {0} mm supera barra comercial (12 m). "
                u"El traslape sigue disponible."
            ).format(format_mm_es(main))

    def _clear_canvas(self, canvas):
        if canvas is None:
            return
        try:
            canvas.Children.Clear()
        except Exception:
            pass

    def _add_line(self, canvas, x1, y1, x2, y2, color, thick=1.5, dash=None):
        ln = Line()
        ln.X1 = float(x1)
        ln.Y1 = float(y1)
        ln.X2 = float(x2)
        ln.Y2 = float(y2)
        ln.Stroke = _brush(color)
        ln.StrokeThickness = float(thick)
        if dash:
            try:
                from System.Windows.Media import DoubleCollection

                dc = DoubleCollection()
                for v in dash:
                    dc.Add(float(v))
                ln.StrokeDashArray = dc
            except Exception:
                pass
        canvas.Children.Add(ln)

    def _add_rect(self, canvas, x, y, w, h, fill, stroke=None, stroke_w=1.0):
        r = Rectangle()
        r.Width = max(1.0, float(w))
        r.Height = max(1.0, float(h))
        if isinstance(fill, SolidColorBrush):
            r.Fill = fill
        else:
            r.Fill = _brush(fill)
        if stroke:
            if isinstance(stroke, SolidColorBrush):
                r.Stroke = stroke
            else:
                r.Stroke = _brush(stroke)
            r.StrokeThickness = float(stroke_w)
        Canvas.SetLeft(r, float(x))
        Canvas.SetTop(r, float(y))
        canvas.Children.Add(r)

    def _add_text(self, canvas, x, y, text, color, size=10):
        tb = TextBlock()
        tb.Text = _as_unicode(text)
        tb.Foreground = _brush(color)
        tb.FontSize = float(size)
        Canvas.SetLeft(tb, float(x))
        Canvas.SetTop(tb, float(y))
        canvas.Children.Add(tb)

    def _add_ellipse(self, canvas, cx, cy, r, fill, stroke, stroke_w=1.2):
        el = Ellipse()
        el.Width = float(r) * 2
        el.Height = float(r) * 2
        el.Fill = fill if isinstance(fill, SolidColorBrush) else _brush(fill)
        el.Stroke = stroke if isinstance(stroke, SolidColorBrush) else _brush(stroke)
        el.StrokeThickness = float(stroke_w)
        Canvas.SetLeft(el, float(cx) - float(r))
        Canvas.SetTop(el, float(cy) - float(r))
        canvas.Children.Add(el)

    def _on_canvas_click(self, sender, args):
        draw = self._elev_draw or {}
        bar_x0 = draw.get(u"bar_x0")
        bar_x1 = draw.get(u"bar_x1")
        main = self._main_mm()
        if bar_x0 is None or bar_x1 is None or main <= 1.0:
            return
        try:
            pos = args.GetPosition(self._canvas)
            x = float(pos.X)
            y = float(pos.Y)
        except Exception:
            return
        band_top = draw.get(u"band_top", 0)
        band_bot = draw.get(u"band_bot", 9999)
        if y < band_top - 12 or y > band_bot + 20:
            return
        span = float(bar_x1) - float(bar_x0)
        if abs(span) < 1.0:
            return
        t = (x - float(bar_x0)) / span
        t = max(0.0, min(1.0, t))
        mm = t * main
        self._cuts_mm = toggle_cut_at_mm(self._cuts_mm, mm, main)
        self._redraw()
        self._set_status(self._status_line())

    def _redraw(self):
        c = self._canvas
        if c is None:
            return
        try:
            self._redraw_elevation_canvas(c)
        except Exception as ex:
            self._elev_draw = {}
            try:
                self._set_status(
                    u"Error vista elevación: {0}".format(_as_unicode(ex))
                )
            except Exception:
                pass

    def _neighbor_elev_height_px(self, neighbor_items, host_h_mm, px_v, wall_top_y, wall_bot_y):
        """Alto en px del vecino en elevación (alineado al pie del muro host)."""
        host_px = max(2.0, abs(float(wall_bot_y) - float(wall_top_y)))
        best_h_mm = float(host_h_mm or 0.0)
        for item in neighbor_items or []:
            try:
                best_h_mm = max(best_h_mm, float(item.get(u"height_mm") or 0.0))
            except Exception:
                pass
        if best_h_mm <= 1.0:
            best_h_mm = float(host_h_mm or 2800.0)
        nh_px = max(2.0, best_h_mm * px_v)
        nh_px = min(nh_px, host_px * 1.35)
        ny = float(wall_bot_y) - nh_px
        return ny, nh_px

    def _draw_elev_end_neighbors(
        self,
        canvas,
        x_span,
        end_items,
        mm_inner,
        mm_outer,
        wall_top_y,
        wall_bot_y,
        host_h_mm,
        px_v,
        fund_top_y=None,
        fund_h_px=0.0,
    ):
        """
        Dibuja muro(s) vecino(s) en un extremo: ancho = espesor real (mm → px_h vía x_span).
        """
        if not end_items:
            return
        try:
            width_mm = max(float(n.get(u"width_mm") or 0.0) for n in end_items)
        except Exception:
            width_mm = 0.0
        if width_mm <= 1.0:
            return
        nx, nw = x_span(float(mm_outer), float(mm_inner))
        nw = max(2.0, nw)
        ny, nh = self._neighbor_elev_height_px(
            end_items, host_h_mm, px_v, wall_top_y, wall_bot_y
        )
        self._add_rect(
            canvas,
            nx,
            ny,
            nw,
            nh,
            _brush(_JOINED, 120),
            _JOINED_STROKE,
            1.2,
        )
        if fund_top_y is not None and fund_h_px > 2.0:
            self._add_rect(
                canvas,
                nx,
                float(fund_top_y),
                nw,
                max(2.0, float(fund_h_px)),
                _brush(_JOINED, 90),
                _JOINED_STROKE,
                1.0,
            )
        if nw > 18.0:
            self._add_text(
                canvas,
                nx + max(2.0, nw * 0.5 - 16.0),
                ny + nh * 0.42,
                u"e={0:.0f}".format(width_mm),
                u"#64748b",
                8,
            )

    def _redraw_elevation_canvas(self, c):
        """
        Elevación a escala: muro, fundación, muros unidos en extremos,
        tramo mayor (margen ± estirón), empotramiento y pata con cotas reales.
        """
        self._clear_canvas(c)
        w = float(c.ActualWidth or 0)
        h = float(c.ActualHeight or 0)
        if w < 40 or h < 40:
            self._elev_draw = {}
            return

        meta = self._meta or {}
        L = max(1.0, float(meta.get(u"length_mm") or 1000.0))
        H = max(1.0, float(meta.get(u"height_mm") or 2800.0))
        fund = meta.get(u"foundation")
        joined = meta.get(u"joined")
        main = max(1.0, self._main_mm())
        n_capas = len(self._layers)
        stretch_a = max(0.0, float(meta.get(u"join_stretch_start_mm") or 0.0))
        stretch_b = max(0.0, float(meta.get(u"join_stretch_end_mm") or 0.0))
        neigh_w_a = max(0.0, float(meta.get(u"end_neighbor_width_inicio_mm") or 0.0))
        neigh_w_b = max(0.0, float(meta.get(u"end_neighbor_width_fin_mm") or 0.0))
        if neigh_w_a <= 1.0 and stretch_a > 1.0:
            neigh_w_a = stretch_a * 2.0
        if neigh_w_b <= 1.0 and stretch_b > 1.0:
            neigh_w_b = stretch_b * 2.0
        end_neigh = meta.get(u"end_neighbors") or {}

        superior = self._is_superior_mode()
        fund_h_mm = float(fund[u"height_mm"]) if fund else 0.0
        margin = float(MARGIN_END_MM)

        # Extremos del segmento mayor (sin empotramiento), mm desde P0
        clear0 = margin - stretch_a
        clear1 = L - margin + stretch_b

        grade = self._concrete_grade
        embed_a_max = 0.0
        embed_b_max = 0.0
        for ly in self._layers or [{u"diam_mm": 10}]:
            d = float(ly.get(u"diam_mm", 10))
            emb = float(empotramiento_mm_from_diam(d, grade))
            if self._end_a == END_EMPOTRO:
                embed_a_max = max(embed_a_max, emb)
            if self._end_b == END_EMPOTRO:
                embed_b_max = max(embed_b_max, emb)

        bar0 = clear0 - (embed_a_max if self._end_a == END_EMPOTRO else 0.0)
        bar1 = clear1 + (embed_b_max if self._end_b == END_EMPOTRO else 0.0)

        # Mundo horizontal: muros unidos a cara exterior del extremo
        x_lo = min(0.0, -neigh_w_a, bar0)
        x_hi = max(L, L + neigh_w_b, bar1)
        z_bot = 0.0
        z_fund_top = fund_h_mm if fund else 0.0
        z_wall_bot = z_fund_top
        z_wall_top = z_wall_bot + H
        z_hi = z_wall_top
        if joined and joined.get(u"relation") == u"stacked_above":
            z_hi += min(600.0, H * 0.25)

        pad_l = 28.0
        pad_r = 20.0
        pad_t = 18.0
        pad_b = 22.0
        usable_w = max(40.0, w - pad_l - pad_r)
        usable_h = max(40.0, h - pad_t - pad_b)
        content_w = max(1.0, x_hi - x_lo)
        content_h = max(1.0, z_hi - z_bot + 40.0)
        # Escala horizontal = largo modelado fiel; vertical solo para encajar alzado
        px_h = max(0.02, usable_w / content_w)
        px_v = max(0.02, usable_h / content_h)

        flip = self._elev_flip_flag()

        def X(mm):
            # mm desde P0 (LocationCurve). Con flip: izquierda = mayor mm (P1).
            if flip:
                return pad_l + (x_hi - float(mm)) * px_h
            return pad_l + (float(mm) - x_lo) * px_h

        def Y(z_mm):
            return (h - pad_b) - (float(z_mm) - z_bot) * px_v

        def x_span(mm0, mm1):
            a = X(mm0)
            b = X(mm1)
            return min(a, b), abs(b - a)

        ground_y = Y(0.0)
        fund_top_y = Y(z_fund_top)
        wall_bot_y = Y(z_wall_bot)
        wall_top_y = Y(z_wall_top)
        wall_x, wall_w = x_span(0.0, L)
        wall_w = max(2.0, wall_w)
        wall_h = max(2.0, H * px_v)

        self._add_line(c, 8, ground_y, w - 8, ground_y, u"#21465C", 1.0, (4, 3))

        # Fundación a lo largo del muro (+ solape bajo vecinos)
        if fund:
            fund_x0 = min(0.0, -neigh_w_a * 0.5) if neigh_w_a > 0 else 0.0
            fund_x1 = max(L, L + neigh_w_b * 0.5) if neigh_w_b > 0 else L
            fx, fw = x_span(fund_x0, fund_x1)
            fw = max(2.0, fw)
            fh = max(2.0, fund_h_mm * px_v)
            self._add_rect(
                c, fx, fund_top_y, fw, fh, _brush(_FUND), _FUND_STROKE, 1.5
            )
            if not superior:
                self._add_rect(
                    c,
                    fx + 2,
                    fund_top_y + 2,
                    max(4.0, fw - 4),
                    max(4.0, fh - 4),
                    _brush(_HOST_STROKE, 40),
                    _HOST_STROKE,
                    1.2,
                )
                self._add_text(
                    c, fx + 6, fund_top_y + 8, u"HOST · WallFoundation", _HOST_STROKE, 9
                )
            self._add_text(
                c,
                fx + fw * 0.5 - 36,
                fund_top_y + fh * 0.5 - 4,
                u"h={0:.0f} mm".format(fund_h_mm),
                u"#95B8CC",
                9,
            )

        # Muros unidos en extremos (ancho = espesor real del vecino, escala px_h)
        fund_h_px = max(0.0, float(ground_y) - float(fund_top_y)) if fund else 0.0
        self._draw_elev_end_neighbors(
            c,
            x_span,
            end_neigh.get(u"inicio") or [],
            0.0,
            -neigh_w_a,
            wall_top_y,
            wall_bot_y,
            H,
            px_v,
            fund_top_y=fund_top_y if fund else None,
            fund_h_px=fund_h_px if fund else 0.0,
        )
        self._draw_elev_end_neighbors(
            c,
            x_span,
            end_neigh.get(u"fin") or [],
            L,
            L + neigh_w_b,
            wall_top_y,
            wall_bot_y,
            H,
            px_v,
            fund_top_y=fund_top_y if fund else None,
            fund_h_px=fund_h_px if fund else 0.0,
        )

        if joined and joined.get(u"relation") == u"stacked_above":
            jh_mm = min(600.0, H * 0.25)
            self._add_rect(
                c,
                wall_x + 4,
                Y(z_wall_top + jh_mm),
                max(4.0, wall_w - 8),
                max(4.0, jh_mm * px_v - 2),
                _brush(_JOINED),
                _JOINED_STROKE,
                1.0,
            )
            self._add_text(
                c,
                wall_x + wall_w * 0.5 - 28,
                Y(z_wall_top + jh_mm * 0.55),
                u"unido ↑",
                u"#64748b",
                9,
            )

        self._add_rect(
            c, wall_x, wall_top_y, wall_w, wall_h, _brush(_WALL), _WALL_STROKE, 1.5
        )
        if superior:
            band_h = max(6.0, min(22.0, COVER_SUPERIOR_MM * px_v * 2.5))
            self._add_rect(
                c,
                wall_x + 2,
                wall_top_y + 2,
                max(4.0, wall_w - 4),
                band_h,
                _brush(_ACCENT, 35),
                _ACCENT,
                1.2,
            )
            self._add_text(c, wall_x + 8, wall_top_y + 8, u"HOST · Muro", _ACCENT, 9)

        try:
            wid = meta.get(u"id")
            self._add_text(
                c,
                wall_x + wall_w * 0.5 - 54,
                wall_top_y + wall_h * 0.42,
                u"Muro {0} · L={1:.2f} m".format(wid, L / 1000.0),
                u"#E8F4F8",
                10,
            )
        except Exception:
            pass

        if joined and joined.get(u"relation") == u"stacked_below":
            self._add_text(
                c,
                wall_x + wall_w * 0.5 - 24,
                fund_top_y - 12 if fund else ground_y - 12,
                u"unido ↓",
                u"#64748b",
                9,
            )

        # Guías de margen 50 mm
        self._add_line(
            c, X(margin), wall_bot_y - 2, X(margin), wall_top_y + 2, u"#21465C", 0.8, (2, 2)
        )
        self._add_line(
            c,
            X(L - margin),
            wall_bot_y - 2,
            X(L - margin),
            wall_top_y + 2,
            u"#21465C",
            0.8,
            (2, 2),
        )

        # Segmento mayor = mismo vano que CreateFromCurves (antes de empotramiento)
        clear_mm = max(1.0, clear1 - clear0)
        # clear_x0/x1 = pixeles extremos A→B (pueden ir der→izq si flip)
        clear_x0 = X(clear0)
        clear_x1 = X(clear1)

        base_cover = (
            COVER_SUPERIOR_MM
            if superior
            else (COVER_FUND_BOT_MM if fund else COVER_WALL_BOT_MM)
        )
        band_ys = []
        for li in range(max(1, n_capas)):
            off = cover_axis_offset_mm_for_layer(
                self._layers, li, base_cover_mm=base_cover
            )
            if superior:
                z_bar = z_wall_top - off
            else:
                z_bar = z_bot + off
            band_ys.append(Y(z_bar))

        if band_ys:
            band_top = min(band_ys) - 14
            band_bot = max(band_ys) + 14
        else:
            band_top = wall_bot_y - 20
            band_bot = wall_bot_y + 20

        self._elev_draw = {
            u"bar_x0": float(clear_x0),
            u"bar_x1": float(clear_x1),
            u"band_top": float(band_top),
            u"band_bot": float(band_bot),
        }

        bar_span = float(clear_x1) - float(clear_x0)  # signed if flip

        def x_at(mm):
            # Cortes UI en 0..main (A→B) orientados a la vista activa
            return clear_x0 + (float(mm) / max(1.0, main)) * bar_span

        ref_horiz = 0.0
        ref_devel = 0.0
        for li, ly in enumerate(self._layers):
            d = float(ly.get(u"diam_mm", 10))
            color = _LAYER_COLORS[li % len(_LAYER_COLORS)]
            lap = traslape_mm_from_diam(d, grade)
            cuts = stagger_cuts_for_layer(self._cuts_mm, li, main, lap)
            embed_a = (
                float(empotramiento_mm_from_diam(d, grade))
                if self._end_a == END_EMPOTRO
                else 0.0
            )
            embed_b = (
                float(empotramiento_mm_from_diam(d, grade))
                if self._end_b == END_EMPOTRO
                else 0.0
            )
            pata_a = (
                float(pata_mm_from_diam(d, concrete_grade=grade))
                if self._end_a == END_PATA
                else 0.0
            )
            pata_b = (
                float(pata_mm_from_diam(d, concrete_grade=grade))
                if self._end_b == END_PATA
                else 0.0
            )
            # Largos idénticos a _build_longitudinal_curves
            horiz_mm = clear_mm + embed_a + embed_b
            devel_mm = horiz_mm + pata_a + pata_b
            if li == 0:
                ref_horiz = horiz_mm
                ref_devel = devel_mm
            stroke = max(1.4, min(3.2, d * px_h * 0.45))

            off = cover_axis_offset_mm_for_layer(
                self._layers, li, base_cover_mm=base_cover
            )
            if superior:
                y = Y(z_wall_top - off)
                pata_sign = -1.0
            else:
                y = Y(z_bot + off)
                pata_sign = 1.0

            for _ci, cut in enumerate(cuts):
                a, b = lap_zone_around_cut(cut, lap, self._lap_mode)
                a = max(0.0, a)
                b = min(main, b)
                xa, xb = x_at(a), x_at(b)
                self._add_rect(
                    c,
                    min(xa, xb),
                    y - stroke - 2,
                    max(2.0, abs(xb - xa)),
                    stroke + 6,
                    _brush(color, 70),
                )

            # Extremos horizontales: tramo mayor ± empotramientos (mm mundo → X vista)
            x0 = X(clear0 - embed_a)
            x1 = X(clear1 + embed_b)

            if cuts:
                pts = [0.0] + list(cuts) + [main]
                for pi in range(len(pts) - 1):
                    stroke_c = _TRAMO_COLORS[pi % len(_TRAMO_COLORS)]
                    xa = x_at(pts[pi])
                    xb = x_at(pts[pi + 1])
                    if pi == 0 and embed_a > 1e-6:
                        xa = x0
                    if pi == len(pts) - 2 and embed_b > 1e-6:
                        xb = x1
                    self._add_line(c, xa, y, xb, y, stroke_c, stroke)
            else:
                self._add_line(c, x0, y, x1, y, color, stroke)

            # Patas L: longitud de eje real en vertical (px_v)
            xa_end = X(clear0)
            xb_end = X(clear1)
            if pata_a > 1.0:
                dy = -pata_sign * pata_a * px_v
                self._add_line(c, xa_end, y, xa_end, y + dy, color, stroke)
            if pata_b > 1.0:
                dy = -pata_sign * pata_b * px_v
                self._add_line(c, xb_end, y, xb_end, y + dy, color, stroke)

            for ci, cut in enumerate(cuts):
                self._add_line(
                    c, x_at(cut), y - 10, x_at(cut), y + 10, color, _STROKE_CUT, (3, 2)
                )
                self._add_ellipse(c, x_at(cut), y, 3.5 if li else 4.5, u"#050E18", color)
                if li == 0:
                    self._add_text(
                        c, x_at(cut) - 10, y - 16, u"C1-{0}".format(ci + 1), _CUT, 8
                    )

            if li == 0:
                mid_x = 0.5 * (x0 + x1)
                self._add_text(
                    c,
                    mid_x - 70,
                    y - 18,
                    u"L={0} mm".format(format_mm_es(horiz_mm)),
                    color,
                    9,
                )

        # Etiquetas A/B en extremos LocationCurve (orientadas a la vista)
        self._add_text(
            c,
            X(0.0) + (-10.0 if not flip else 4.0),
            wall_top_y + 4,
            u"A",
            u"#fbbf24",
            11,
        )
        self._add_text(
            c,
            X(L) + (-4.0 if not flip else -10.0),
            wall_top_y + 4,
            u"B",
            u"#fbbf24",
            11,
        )

        foot = u"L modelada {0} mm".format(format_mm_es(ref_horiz))
        if ref_devel > ref_horiz + 1.0:
            foot += u" · desar. {0} mm".format(format_mm_es(ref_devel))
        foot += u" · tramo {0} mm".format(format_mm_es(main))
        if stretch_a > 1.0 or stretch_b > 1.0:
            foot += u" · estirón +{0:.0f}/+{1:.0f}".format(stretch_a, stretch_b)
        foot += u" · vista {0}".format(u"A→derecha" if flip else u"A→izquierda")
        self._add_text(c, 12, h - 14, foot, u"#64748b", 9)

        pick = self._win.FindName(u"TxtPickHint") if self._win else None
        if pick is not None:
            if not self._cuts_mm:
                pick.Text = (
                    u"Clic en la elevación para ubicar empalmes. "
                    u"Capa 1 es la referencia; las capas 2–3 alternan A-B-A."
                    if n_capas > 1
                    else u"Clic en la elevación para ubicar el empalme sobre el tramo."
                )
            else:
                pick.Text = (
                    u"{0} corte(s) en capa 1. Clic añade otro · clic cerca de C1-# lo quita."
                    .format(len(self._cuts_mm))
                    if n_capas > 1
                    else u"{0} corte(s). Clic añade otro · clic cerca de C# lo quita.".format(
                        len(self._cuts_mm)
                    )
                )

    def _redraw_section(self):
        c = self._canvas_sec
        if c is None:
            return
        self._clear_canvas(c)
        w = float(c.ActualWidth or 0)
        h = float(c.ActualHeight or 0)
        if w < 40 or h < 40:
            return

        meta = self._meta or {}
        thick = float(meta.get(u"thickness_mm") or 200.0)
        height_mm = min(float(meta.get(u"height_mm") or 2800.0), 1200.0)
        fund = meta.get(u"foundation")
        px = min(0.35, max(0.12, (w - 80.0) / max(thick + (fund[u"width_mm"] if fund else 80), 1.0)))

        wall_t = thick * px
        wall_h = min(height_mm * px, 70.0)
        fund_w = (float(fund[u"width_mm"]) * px) if fund else 0.0
        fund_h = (float(fund[u"height_mm"]) * px) if fund else 0.0

        ground_y = h - 22.0
        fund_top = ground_y - fund_h if fund else ground_y
        wall_bot = fund_top
        wall_top = wall_bot - wall_h
        # Muro centrado; fundación centrada bajo el espesor (voladizos fuera del muro).
        wall_x = (w - wall_t) * 0.5
        fund_x = wall_x + (wall_t - fund_w) * 0.5 if fund else wall_x

        self._add_line(c, 10, ground_y, w - 10, ground_y, u"#21465C", 1.0, (3, 2))
        superior = self._is_superior_mode()
        if fund:
            self._add_rect(
                c, fund_x, fund_top, fund_w, fund_h, _brush(_FUND), _FUND_STROKE, 1.5
            )
            if not superior:
                self._add_rect(
                    c,
                    fund_x + 2,
                    fund_top + 2,
                    max(4.0, fund_w - 4),
                    max(4.0, fund_h - 4),
                    _brush(_HOST_STROKE, 35),
                    _HOST_STROKE,
                    1.0,
                )
                self._add_text(c, fund_x + 4, fund_top + 10, u"HOST", _HOST_STROKE, 8)
        self._add_rect(
            c, wall_x, wall_top, wall_t, wall_h, _brush(_WALL), _WALL_STROKE, 1.5
        )
        if superior:
            self._add_rect(
                c,
                wall_x + 2,
                wall_top + 2,
                max(4.0, wall_t - 4),
                min(14.0, max(6.0, wall_h * 0.15)),
                _brush(_ACCENT, 35),
                _ACCENT,
                1.0,
            )
            self._add_text(c, wall_x + 4, wall_top + 10, u"HOST", _ACCENT, 8)
        self._add_text(
            c, wall_x + wall_t * 0.5 - 16, wall_top + wall_h * 0.4, u"e={0}".format(int(thick)), u"#E8F4F8", 9
        )

        cover_bot = COVER_FUND_BOT_MM if fund else COVER_WALL_BOT_MM
        cover_px = cover_bot * px
        max_r = 4.0
        layer_gap = max(max_r * 2 + 4, 12.0)
        if superior:
            cover_top_px = COVER_SUPERIOR_MM * px
            base_y = wall_top + cover_top_px + max_r
        elif fund:
            base_y = ground_y - cover_px - max_r
        else:
            base_y = wall_bot - cover_px - max_r

        # Puntos solo en banda de espesor muro (cover…e−cover), nunca en voladizo.
        for li, ly in enumerate(self._layers):
            offsets = layer_bar_offsets_mm(thick, ly.get(u"n_bars", 2), COVER_WALL_MM)
            cy = base_y + li * layer_gap if superior else base_y - li * layer_gap
            color = _LAYER_COLORS[li % len(_LAYER_COLORS)]
            r = max(2.5, float(ly.get(u"diam_mm", 10)) * px * 0.35)
            for ox in offsets:
                cx = wall_x + ox * px
                self._add_ellipse(c, cx, cy, r, color, color, 0.5)

        self._add_text(
            c,
            8,
            6,
            u"Sección · {0} capa(s) · espesor muro · host {1}".format(
                len(self._layers),
                u"muro" if superior else u"fundación",
            ),
            u"#64748b",
            9,
        )

    def show(self):
        self._win.Show()

    def _execute_colocar(self):
        self._read_layer_combos()
        self._concrete_grade = self._read_concrete_grade()
        self._end_a = self._read_end(self._cmb_end_a, END_PATA)
        self._end_b = self._read_end(self._cmb_end_b, END_PATA)
        self._mode = self._read_mode()
        res = place_retorno_malla_fundacion(
            self._doc,
            self._uidoc,
            self._walls,
            self._layers,
            cuts_ref_mm=list(self._cuts_mm),
            lap_mode_ui=self._lap_mode,
            concrete_grade=self._concrete_grade,
            end_a=self._end_a,
            end_b=self._end_b,
            mode=self._mode,
        )
        if res.get(u"ok"):
            try:
                self._win.Close()
            except Exception:
                pass
            return
        msgs = u"\n".join(res.get(u"messages") or [])
        _mostrar_aviso(
            self._uiapp,
            u"No se pudo colocar la armadura.",
            msgs or u"Error desconocido.",
        )
        self._set_status(u"Error al colocar.")


def show_retorno_malla_fundacion_window(uiapp, uidoc, doc, walls, on_closed=None):
    if _focus_existing():
        return None
    try:
        win = RetornoMallaHostFundacionWindow(
            uiapp, uidoc, doc, walls, on_closed=on_closed
        )
        win.show()
        return win
    except Exception:
        _clear_singleton()
        raise
