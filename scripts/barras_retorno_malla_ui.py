# -*- coding: utf-8 -*-
"""
UI WPF — Arainco: Barras de retorno de malla.

Elevación izquierda · sección + parámetros a la derecha (como mockup validado).
Dosificación G25/G35/G45, capas n/Ø, Ext. A/B (Pata 90º | Empotramiento),
empalme por pick con alternancia A-B-A (Coronamiento / 56).

No muestra cards de estado MURO UNIDO / FUNDACIÓN; la detección es interna.
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
    END_OPTS,
    END_PATA,
    LAP_MODE_LABELS,
    LAP_MODE_SYMMETRIC,
    LAYER_CENTERLINE_SPACING_MM,
    MAX_BARRA_COMERCIAL_MM,
    N_BARS_OPTS,
    clamp_n_capas,
    end_label,
    format_mm_es,
    lap_zone_around_cut,
    layer_bar_offsets_mm,
    normalize_concrete_grade,
    normalize_end_condition,
    normalize_lap_mode_ui,
    stagger_cuts_for_layer,
    sync_layers,
    toggle_cut_at_mm,
    traslape_mm_from_diam,
)
from barras_retorno_malla_place import (
    place_barras_retorno_malla,
    wall_meta_for_ui,
)
from revit_wpf_window_position import revit_main_hwnd

_DIALOG_TITLE = u"Arainco: Barras de retorno de malla"
_SINGLETON_KEY = u"Arainco.BarrasRetornoMalla.ActiveWindow"
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

_UI_WIDTH = 1040
_UI_MIN_WIDTH = 960
_UI_SIDE_W = 360
_UI_CANVAS_H = 400
_UI_SECTION_H = 180
_STROKE_BAR = 1.6
_STROKE_WALL = 1.5
_STROKE_CUT = 1.4


def _ui_min_window_height():
    chrome = 200
    stack = 28 + 40 + 22 + 54 + 100 + 90 + 140
    body = max(_UI_CANVAS_H + 34, _UI_SECTION_H + stack) + 8
    return int(chrome + body)


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


def _focus_existing():
    try:
        win = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        win = None
    if win is None:
        return False
    try:
        if not bool(win.IsLoaded):
            return False
    except Exception:
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


class _ColocarHandler(IExternalEventHandler):
    def __init__(self, win_ref):
        self._win_ref = win_ref

    def GetName(self):
        return u"AraincoBarrasRetornoMallaColocar"

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


class BarrasRetornoMallaWindow(object):
    def __init__(self, uiapp, uidoc, doc, walls, on_closed=None):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        self._walls = list(walls or [])
        self._wall = self._walls[0] if self._walls else None
        self._on_closed = on_closed
        self._meta = wall_meta_for_ui(doc, self._wall) if self._wall else {}
        self._layers = sync_layers(None, 1)
        self._n_capas = 1
        self._lap_mode = LAP_MODE_SYMMETRIC
        self._concrete_grade = DOSIFICACION_HORMIGON_DEFAULT
        self._end_a = END_EMPOTRO
        self._end_b = END_PATA
        self._cuts_mm = []
        self._cmb_capas = None
        self._cmb_lap = None
        self._cmb_dosif = None
        self._cmb_end_a = None
        self._cmb_end_b = None
        self._layer_host = None
        self._elev_draw = {}
        self._canvas = None
        self._canvas_sec = None
        self._txt_status = None
        self._txt_warn = None
        self._warn_border = None
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
        hint = u"Pick empalme en elevación · extremos A/B · dosificación G"
        min_h = _ui_min_window_height()
        xaml = build_simple_tool_xaml(
            title=_DIALOG_TITLE,
            styles_xml=BIMTOOLS_DARK_STYLES_XML,
            body_xaml=body,
            footer_actions_xaml=footer,
            footer_hint_xaml=hint,
            width=_UI_WIDTH,
            min_width=_UI_MIN_WIDTH,
            min_height=min_h,
            resize_mode=u"CanResizeWithGrip",
            size_to_content_height=True,
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
        try:
            self._win.Width = float(_UI_WIDTH)
            self._win.MinWidth = float(_UI_MIN_WIDTH)
            self._win.MinHeight = float(min_h)
        except Exception:
            pass
        self._wire()
        self._rebuild_layer_cards()
        self._redraw()
        self._redraw_section()
        self._refresh_warn()
        self._set_status(self._status_line())
        try:
            self._win.Loaded += RoutedEventHandler(self._on_loaded_fit)
        except Exception:
            pass

        def _on_closed(sender, args):
            self._closed = True
            try:
                AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
            except Exception:
                pass
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

    def _main_mm(self):
        try:
            return float(self._meta.get(u"main_mm") or 0.0)
        except Exception:
            return 0.0

    def _build_body_xaml(self):
        xaml = u"""
<Grid MinHeight="__CANVAS_H__">
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="*"/>
    <ColumnDefinition Width="__SIDE_W__"/>
  </Grid.ColumnDefinitions>
  <Border Grid.Column="0" Background="#050E18" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" Margin="0,0,10,0" MinHeight="__CANVAS_H__">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="__CANVAS_H__"/>
      </Grid.RowDefinitions>
      <DockPanel Grid.Row="0" Margin="8,6">
        <TextBlock x:Name="TxtElevHdr" DockPanel.Dock="Left"
                   Text="ELEVACIÓN · MURO"
                   Foreground="#95B8CC" FontSize="11" VerticalAlignment="Center"/>
        <Border DockPanel.Dock="Right" BorderBrush="#5BC0DE" BorderThickness="1"
                Background="#0E1B32" Padding="6,2" CornerRadius="3">
          <TextBlock Text="LONG." Foreground="#5BC0DE" FontSize="10" FontWeight="SemiBold"/>
        </Border>
      </DockPanel>
      <TextBlock Grid.Row="1" x:Name="TxtPickHint" Margin="8,0,8,4"
                 Foreground="#fbbf24" FontSize="10" TextWrapping="Wrap"
                 Text="Clic en elevación: añade empalme · clic cerca: quita"/>
      <Canvas x:Name="CnvElev" Grid.Row="2" Background="#050E18" ClipToBounds="True"
              Height="__CANVAS_H__" MinHeight="__CANVAS_H__" Cursor="Cross"/>
    </Grid>
  </Border>
  <Border x:Name="PanelStack" Grid.Column="1" Background="#0a1620" BorderBrush="#5BC0DE"
          BorderThickness="1.5" CornerRadius="4" Padding="10" VerticalAlignment="Top">
    <StackPanel>
        <TextBlock Text="SECCIÓN" Foreground="#64748b" FontSize="10" FontWeight="SemiBold"
                   Margin="0,0,0,4"/>
        <Border BorderBrush="#21465C" BorderThickness="1" CornerRadius="4" Margin="0,0,0,10">
          <Canvas x:Name="CnvSec" Background="#0E1B32" Height="__SEC_H__" ClipToBounds="True"/>
        </Border>
        <DockPanel Margin="0,0,0,6">
          <StackPanel DockPanel.Dock="Right" Orientation="Horizontal"
                      VerticalAlignment="Center">
            <ComboBox x:Name="CmbDosificacionHormigon" Style="{StaticResource Combo}"
                      MinWidth="64" Margin="0,0,6,0" VerticalAlignment="Center"
                      IsEditable="False" IsReadOnly="True"
                      ToolTip="Dosificación del hormigón"/>
            <Border BorderBrush="#5BC0DE" BorderThickness="1"
                    Background="#0E1B32" Padding="6,2" CornerRadius="3">
              <TextBlock Text="LONG." Foreground="#5BC0DE" FontSize="10" FontWeight="SemiBold"/>
            </Border>
          </StackPanel>
          <TextBlock Text="Parámetros" Foreground="#E8F4F8" FontSize="12"
                     FontWeight="SemiBold" VerticalAlignment="Center"/>
        </DockPanel>
        <TextBlock x:Name="TxtParamsHint" Foreground="#64748b" FontSize="10"
                   TextWrapping="Wrap" Margin="0,0,0,8"
                   Text="Barras en espesor muro (≥2) · recub. caras 25 mm"/>
        <TextBlock Text="CAPAS" Foreground="#64748b" FontSize="10" FontWeight="SemiBold"
                   Margin="0,0,0,4"/>
        <ComboBox x:Name="CmbCapas" Style="{StaticResource ComboStretch}"
                  HorizontalAlignment="Stretch" Margin="0,0,0,8"/>
        <StackPanel x:Name="PanelLayers"/>
        <Border Height="1" Background="#21465C" Margin="0,4,0,8"/>
        <TextBlock Text="EXTREMOS" Foreground="#64748b" FontSize="10" FontWeight="SemiBold"
                   Margin="0,0,0,4"/>
        <DockPanel Margin="0,0,0,6">
          <TextBlock Text="Ext. A" Foreground="#95B8CC" FontSize="11" Width="56"
                     VerticalAlignment="Center"/>
          <ComboBox x:Name="CmbEndA" Style="{StaticResource ComboStretch}"
                    HorizontalAlignment="Stretch"/>
        </DockPanel>
        <DockPanel Margin="0,0,0,8">
          <TextBlock Text="Ext. B" Foreground="#95B8CC" FontSize="11" Width="56"
                     VerticalAlignment="Center"/>
          <ComboBox x:Name="CmbEndB" Style="{StaticResource ComboStretch}"
                    HorizontalAlignment="Stretch"/>
        </DockPanel>
        <Border Height="1" Background="#21465C" Margin="0,0,0,8"/>
        <TextBlock Text="EMPALME / SPLIT" Foreground="#64748b" FontSize="10"
                   FontWeight="SemiBold" Margin="0,0,0,4"/>
        <TextBlock x:Name="TxtCutsHint" Foreground="#64748b" FontSize="9"
                   TextWrapping="Wrap" Margin="0,0,0,6"
                   Text="Clic en elevación: añade corte · clic cerca: quita"/>
        <ComboBox x:Name="CmbLap" Style="{StaticResource ComboStretch}"
                  HorizontalAlignment="Stretch" Margin="0,0,0,6"/>
        <Border x:Name="WarnBorder" Background="#0E1B32" BorderBrush="#d97706"
                BorderThickness="1" CornerRadius="4" Padding="6" Margin="0,4,0,0"
                Visibility="Collapsed">
          <TextBlock x:Name="TxtWarn" Foreground="#fcd34d" FontSize="9" TextWrapping="Wrap"/>
        </Border>
    </StackPanel>
  </Border>
</Grid>
"""
        xaml = xaml.replace(u"__CANVAS_H__", u"{0}".format(int(_UI_CANVAS_H)))
        xaml = xaml.replace(u"__SIDE_W__", u"{0}".format(int(_UI_SIDE_W)))
        xaml = xaml.replace(u"__SEC_H__", u"{0}".format(int(_UI_SECTION_H)))
        return xaml.strip()

    def _fit_height_to_content(self):
        try:
            win = self._win
            if win is None:
                return
            from System.Windows import SizeToContent

            floor = float(_ui_min_window_height())
            win.MinHeight = floor
            win.SizeToContent = SizeToContent.Height
            win.UpdateLayout()
            h = float(win.ActualHeight or 0)
            if h < 80.0:
                return
            if h < floor:
                h = floor
            win.SizeToContent = SizeToContent.Manual
            win.Height = h
            win.Width = float(_UI_WIDTH)
        except Exception:
            pass

    def _on_loaded_fit(self, sender, args):
        try:
            self._fit_height_to_content()
            self._redraw()
            self._redraw_section()
        except Exception:
            pass

    def _wire(self):
        w = self._win
        self._canvas = w.FindName(u"CnvElev")
        self._canvas_sec = w.FindName(u"CnvSec")
        self._cmb_capas = w.FindName(u"CmbCapas")
        self._cmb_lap = w.FindName(u"CmbLap")
        self._cmb_dosif = w.FindName(u"CmbDosificacionHormigon")
        self._cmb_end_a = w.FindName(u"CmbEndA")
        self._cmb_end_b = w.FindName(u"CmbEndB")
        self._layer_host = w.FindName(u"PanelLayers")
        self._txt_status = w.FindName(u"TxtStatus")
        self._txt_warn = w.FindName(u"TxtWarn")
        self._warn_border = w.FindName(u"WarnBorder")

        self._update_elev_header()
        self._update_params_hint()

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
            for key, lab in END_OPTS:
                it = ComboBoxItem()
                it.Content = lab
                it.Tag = key
                cmb.Items.Add(it)
                if key == selected_key:
                    cmb.SelectedItem = it
            if cmb.SelectedItem is None and cmb.Items.Count:
                cmb.SelectedIndex = 0

        _fill_end_cmb(self._cmb_end_a, END_EMPOTRO)
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
            self._fit_height_to_content()

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
            self._end_a = self._read_end(self._cmb_end_a, END_EMPOTRO)
            self._end_b = self._read_end(self._cmb_end_b, END_PATA)
            self._redraw()
            self._set_status(self._status_line())

        self._cmb_capas.SelectionChanged += SelectionChangedEventHandler(on_capas)
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

    def _update_elev_header(self):
        hdr = self._win.FindName(u"TxtElevHdr") if self._win else None
        if hdr is None:
            return
        if self._has_fund():
            hdr.Text = u"ELEVACIÓN · MURO + FUNDACIÓN"
        else:
            hdr.Text = u"ELEVACIÓN · MURO (sin fundación)"

    def _update_params_hint(self):
        tip = self._win.FindName(u"TxtParamsHint") if self._win else None
        if tip is None:
            return
        if self._has_fund():
            tip.Text = (
                u"Barras en espesor muro (≥2) · recub. caras {0} mm · "
                u"Z: fondo fund. −{1} mm"
            ).format(int(COVER_WALL_MM), int(COVER_FUND_BOT_MM))
        else:
            tip.Text = (
                u"Barras en espesor muro (≥2) · recub. caras {0} mm · "
                u"Z: base muro +{1} mm"
            ).format(int(COVER_WALL_MM), int(COVER_WALL_BOT_MM))

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
            tip.Text = u"Traslape Ø ≈ {0} mm ({1})".format(
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
            card = Border()
            card.Background = _brush(u"#0a1620")
            card.BorderBrush = _brush(u"#21465C")
            card.BorderThickness = Thickness(1)
            try:
                from System.Windows import CornerRadius

                card.CornerRadius = CornerRadius(4)
            except Exception:
                pass
            card.Padding = Thickness(8)
            card.Margin = Thickness(0, 0, 0, 8)

            sp = StackPanel()
            hdr = TextBlock()
            suffix = u""
            if multi and i == 0:
                suffix = u" · pick"
            elif i % 2 == 1:
                suffix = u" · alt"
            hdr.Text = u"{0}ª C.{1}".format(i + 1, suffix)
            hdr.Foreground = _brush(_LAYER_COLORS[i % len(_LAYER_COLORS)])
            hdr.FontSize = 12
            try:
                from System.Windows import FontWeights

                hdr.FontWeight = FontWeights.SemiBold
            except Exception:
                pass
            sp.Children.Add(hdr)

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
            lbl_n.Text = u"n"
            lbl_n.Foreground = _brush(u"#95B8CC")
            lbl_n.FontSize = 11
            lbl_n.Width = 20
            lbl_n.VerticalAlignment = VerticalAlignment.Center
            cmb_n = _mk_combo(N_BARS_OPTS, ly.get(u"n_bars", 2), 56)
            lbl_d = TextBlock()
            lbl_d.Text = u"Ø"
            lbl_d.Foreground = _brush(u"#95B8CC")
            lbl_d.FontSize = 11
            lbl_d.Width = 20
            lbl_d.VerticalAlignment = VerticalAlignment.Center
            cmb_d = _mk_combo(DIAMS_MM, ly.get(u"diam_mm", 10), 72, diam=True)

            tip = TextBlock()
            tip.Foreground = _brush(_LAYER_COLORS[i % len(_LAYER_COLORS)])
            tip.FontSize = 8
            tip.Margin = Thickness(0, 4, 0, 0)
            tip.Opacity = 0.85

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
                u"C{0}:Ø{1}×{2}".format(
                    i + 1, ly.get(u"diam_mm", 10), ly.get(u"n_bars", 2)
                )
            )
        return u"{0} · {1}C · {2} · A={3} · B={4} · {5} corte(s)".format(
            wid,
            len(self._layers),
            u" · ".join(parts),
            end_label(self._end_a),
            end_label(self._end_b),
            len(self._cuts_mm),
        )

    def _set_status(self, text):
        if self._txt_status is not None:
            try:
                self._txt_status.Text = _as_unicode(text)
            except Exception:
                pass

    def _refresh_warn(self):
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
        span = max(1.0, float(bar_x1) - float(bar_x0))
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
        self._clear_canvas(c)
        w = float(c.ActualWidth or 0)
        h = float(c.ActualHeight or 0)
        if w < 40 or h < 40:
            self._elev_draw = {}
            return

        meta = self._meta or {}
        length_mm = float(meta.get(u"length_mm") or 1000.0)
        height_mm = float(meta.get(u"height_mm") or 2800.0)
        fund = meta.get(u"foundation")
        joined = meta.get(u"joined")
        main = self._main_mm()
        n_capas = len(self._layers)

        margin_x = 40.0
        usable_w = max(40.0, w - margin_x * 2)
        px_per_mm_h = usable_w / max(length_mm, 1.0)
        # escala vertical compacta
        fund_h_mm = float(fund[u"height_mm"]) if fund else 0.0
        total_v = height_mm + fund_h_mm + 80.0
        px_per_mm_v = min(0.12, max(0.05, (h - 60.0) / max(total_v, 1.0)))

        ground_y = h - 28.0
        fund_h = fund_h_mm * px_per_mm_v
        wall_h = height_mm * px_per_mm_v
        fund_top_y = ground_y - fund_h if fund else ground_y
        wall_bot_y = fund_top_y
        wall_top_y = wall_bot_y - wall_h
        wall_x = margin_x
        wall_w = length_mm * px_per_mm_h
        fund_extra = min(28.0, (float(fund[u"width_mm"]) * px_per_mm_h * 0.08) if fund else 0.0)
        fund_x = wall_x - fund_extra
        fund_w = wall_w + fund_extra * 2.0

        # terreno
        self._add_line(c, 12, ground_y, w - 12, ground_y, u"#21465C", 1.0, (4, 3))

        if joined and joined.get(u"relation") == u"stacked_above":
            jh = 36.0
            self._add_rect(
                c,
                wall_x + 6,
                wall_top_y - jh,
                wall_w - 12,
                jh - 4,
                _brush(_JOINED),
                _JOINED_STROKE,
                1.0,
            )
            self._add_text(
                c,
                wall_x + wall_w * 0.5 - 30,
                wall_top_y - jh * 0.5 - 4,
                u"unido ↑",
                u"#64748b",
                9,
            )

        if fund:
            self._add_rect(
                c, fund_x, fund_top_y, fund_w, fund_h, _brush(_FUND), _FUND_STROKE, 1.5
            )
            self._add_text(
                c,
                fund_x + fund_w * 0.5 - 40,
                fund_top_y + fund_h * 0.5 - 4,
                u"Fund. · {0} mm".format(int(fund_h_mm)),
                u"#95B8CC",
                9,
            )

        self._add_rect(
            c, wall_x, wall_top_y, wall_w, wall_h, _brush(_WALL), _WALL_STROKE, 1.5
        )
        try:
            wid = meta.get(u"id")
            self._add_text(
                c,
                wall_x + wall_w * 0.5 - 50,
                wall_top_y + wall_h * 0.45,
                u"Id {0} · L={1:.1f} m".format(wid, length_mm / 1000.0),
                u"#E8F4F8",
                10,
            )
        except Exception:
            pass

        if joined and joined.get(u"relation") == u"stacked_below":
            self._add_text(
                c,
                wall_x + wall_w * 0.5 - 24,
                (fund_top_y - 12) if fund else (ground_y - 12),
                u"unido ↓",
                u"#64748b",
                9,
            )

        cover_bot = COVER_FUND_BOT_MM if fund else COVER_WALL_BOT_MM
        cover_bot_px = cover_bot * px_per_mm_v
        gap = max(10.0, min(16.0, LAYER_CENTERLINE_SPACING_MM * px_per_mm_v * 1.2))
        if fund:
            base_y = ground_y - cover_bot_px - _STROKE_BAR
        else:
            base_y = wall_bot_y - cover_bot_px - _STROKE_BAR

        clear_x0 = wall_x + 50.0 * px_per_mm_h
        clear_x1 = wall_x + wall_w - 50.0 * px_per_mm_h
        bar_span = max(1.0, clear_x1 - clear_x0)

        def x_at(mm):
            return clear_x0 + (float(mm) / max(1.0, main)) * bar_span

        ref_lap = traslape_mm_from_diam(
            self._layers[0].get(u"diam_mm", 10) if self._layers else 10,
            self._concrete_grade,
        )
        embed_px = min(36.0, (ref_lap / 1000.0) * length_mm * px_per_mm_h)
        pata_len = max(12.0, 14.0)

        band_top = base_y - max(0, n_capas - 1) * gap - 10
        band_bot = base_y + 14
        self._elev_draw = {
            u"bar_x0": float(clear_x0),
            u"bar_x1": float(clear_x1),
            u"band_top": float(band_top),
            u"band_bot": float(band_bot),
        }

        for li, ly in enumerate(self._layers):
            y = base_y - li * gap
            color = _LAYER_COLORS[li % len(_LAYER_COLORS)]
            diam = float(ly.get(u"diam_mm", 10))
            lap = traslape_mm_from_diam(diam, self._concrete_grade)
            cuts = stagger_cuts_for_layer(self._cuts_mm, li, main, lap)

            for ci, cut in enumerate(cuts):
                a, b = lap_zone_around_cut(cut, lap, self._lap_mode)
                a = max(0.0, a)
                b = min(main, b)
                xa, xb = x_at(a), x_at(b)
                self._add_rect(
                    c,
                    min(xa, xb),
                    y - 4,
                    max(2.0, abs(xb - xa)),
                    _STROKE_BAR + 8,
                    _brush(color, 70),
                )

            x0 = clear_x0 - embed_px if self._end_a == END_EMPOTRO else clear_x0
            x1 = clear_x1 + embed_px if self._end_b == END_EMPOTRO else clear_x1
            # tramos por cortes
            if cuts:
                pts = [0.0] + list(cuts) + [main]
                for pi in range(len(pts) - 1):
                    stroke = _TRAMO_COLORS[pi % len(_TRAMO_COLORS)]
                    xa = x_at(pts[pi])
                    xb = x_at(pts[pi + 1])
                    if pi == 0 and self._end_a == END_EMPOTRO:
                        xa = x0
                    if pi == len(pts) - 2 and self._end_b == END_EMPOTRO:
                        xb = x1
                    self._add_line(c, xa, y, xb, y, stroke, _STROKE_BAR)
            else:
                self._add_line(c, x0, y, x1, y, color, _STROKE_BAR)

            if self._end_a == END_PATA:
                self._add_line(c, clear_x0, y, clear_x0, y - pata_len, color, _STROKE_BAR)
            if self._end_b == END_PATA:
                self._add_line(c, clear_x1, y, clear_x1, y - pata_len, color, _STROKE_BAR)

            for ci, cut in enumerate(cuts):
                self._add_line(
                    c, x_at(cut), y - 10, x_at(cut), y + 10, color, _STROKE_CUT, (3, 2)
                )
                self._add_ellipse(c, x_at(cut), y, 3.5 if li else 4.5, u"#050E18", color)
                if li == 0:
                    self._add_text(
                        c, x_at(cut) - 10, y - 16, u"C1-{0}".format(ci + 1), _CUT, 8
                    )

        foot = u"{0}C · {1} · A={2} · B={3}".format(
            n_capas,
            self._concrete_grade,
            end_label(self._end_a),
            end_label(self._end_b),
        )
        if fund:
            foot += u" · Z fund −{0}".format(int(COVER_FUND_BOT_MM))
        else:
            foot += u" · Z muro +{0}".format(int(COVER_WALL_BOT_MM))
        self._add_text(c, 12, h - 14, foot, u"#64748b", 9)

        pick = self._win.FindName(u"TxtPickHint") if self._win else None
        if pick is not None:
            if not self._cuts_mm:
                pick.Text = (
                    u"Clic en elevación → empalme capa 1 (ref.); capas 2–3 se alternan"
                    if n_capas > 1
                    else u"Clic en la barra (elevación) para ubicar el empalme"
                )
            else:
                pick.Text = (
                    u"Clic añade corte en capa 1 · clic cerca de C1-# lo quita · A-B-A"
                    if n_capas > 1
                    else u"Clic añade corte · clic cerca de C# lo quita · zona = traslape"
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
        if fund:
            self._add_rect(
                c, fund_x, fund_top, fund_w, fund_h, _brush(_FUND), _FUND_STROKE, 1.5
            )
        self._add_rect(
            c, wall_x, wall_top, wall_t, wall_h, _brush(_WALL), _WALL_STROKE, 1.5
        )
        self._add_text(
            c, wall_x + wall_t * 0.5 - 16, wall_top + wall_h * 0.4, u"e={0}".format(int(thick)), u"#E8F4F8", 9
        )

        cover_bot = COVER_FUND_BOT_MM if fund else COVER_WALL_BOT_MM
        cover_px = cover_bot * px
        max_r = 4.0
        layer_gap = max(max_r * 2 + 4, 12.0)
        if fund:
            base_y = ground_y - cover_px - max_r
        else:
            base_y = wall_bot - cover_px - max_r

        # Puntos solo en banda de espesor muro (cover…e−cover), nunca en voladizo.
        for li, ly in enumerate(self._layers):
            offsets = layer_bar_offsets_mm(thick, ly.get(u"n_bars", 2), COVER_WALL_MM)
            cy = base_y - li * layer_gap
            color = _LAYER_COLORS[li % len(_LAYER_COLORS)]
            r = max(2.5, float(ly.get(u"diam_mm", 10)) * px * 0.35)
            for ox in offsets:
                cx = wall_x + ox * px
                self._add_ellipse(c, cx, cy, r, color, color, 0.5)

        self._add_text(
            c,
            8,
            6,
            u"Sección · {0} capa(s) · en espesor muro{1}".format(
                len(self._layers), u"" if fund else u" · sin fund."
            ),
            u"#64748b",
            9,
        )

    def show(self):
        self._win.Show()

    def _execute_colocar(self):
        self._read_layer_combos()
        self._concrete_grade = self._read_concrete_grade()
        self._end_a = self._read_end(self._cmb_end_a, END_EMPOTRO)
        self._end_b = self._read_end(self._cmb_end_b, END_PATA)
        res = place_barras_retorno_malla(
            self._doc,
            self._uidoc,
            self._walls,
            self._layers,
            cuts_ref_mm=list(self._cuts_mm),
            lap_mode_ui=self._lap_mode,
            concrete_grade=self._concrete_grade,
            end_a=self._end_a,
            end_b=self._end_b,
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


def show_barras_retorno_window(uiapp, uidoc, doc, walls, on_closed=None):
    if _focus_existing():
        return None
    win = BarrasRetornoMallaWindow(
        uiapp, uidoc, doc, walls, on_closed=on_closed
    )
    win.show()
    return win
