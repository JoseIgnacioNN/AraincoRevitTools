# -*- coding: utf-8 -*-
"""
UI WPF — Arainco: Coronamiento muros.

Elevación + stack lateral (capas n/Ø estilo V3), pick de empalme en canvas,
escenario de traslape global (mismos modos que 56_DividirRebarPuntoTraslape),
aviso >12 m (no bloquea).

Modos (multi-pick previo):
  - 1 muro → U libre (U + 2 patas)
  - 2 apilados → Empotrado (V3 voladizo INF en host Z baja)

Elevación: siluetas desde LocationCurve + altura reales proyectadas sobre
``ActiveView.RightDirection`` (X) y Z (Y canvas). Con 2 muros apilados se
conservan offsets relativos (voladizo / empotro). El flip izq./der. solo
espeja el mapeo mm↔píxel de barras/cortes (origen mm = LocationCurve P0).

Flag ``ENABLE_TRASLAPOS``: poner False para ocultar empalme sin quitar el código.
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
from coronamiento_muros_geom import (
    CAPAS_OPTS,
    DIAMS_MM,
    LAP_MODE_LABELS,
    LAP_MODE_SYMMETRIC,
    N_BARS_OPTS,
    clamp_n_capas,
    format_mm_es,
    normalize_lap_mode_ui,
    stagger_cuts_for_layer,
    sync_layers,
    toggle_cut_at_mm,
    traslape_mm_from_diam,
)
from coronamiento_muros_place import (
    place_coronamiento,
    wall_elev_canvas_flip_for_view,
    wall_largo_mm,
    wall_length_estimate,
    wall_length_estimate_empotrado,
    walls_elev_layout_model,
)
from revit_wpf_window_position import revit_main_hwnd

# Poner False para ocultar empalme / no dividir al colocar.
ENABLE_TRASLAPOS = True

_DIALOG_TITLE = u"Arainco: Coronamiento muros"
_SINGLETON_KEY = u"Arainco.CoronamientoMuros.ActiveWindow"
_COR = u"#fbbf24"
_CUT = u"#f87171"
_WALL = u"#2a4a58"
_WALL_STROKE = u"#5a8a9a"
_EMBED = u"#34d399"
_UPPER_FILL = u"#1a3544"
_UPPER_STROKE = u"#7a9aaa"
_LAYER_COLORS = (u"#fbbf24", u"#38bdf8", u"#a78bfa")
_GEOM_U_LIBRE = u"u_libre"
_GEOM_EMPOTRADO = u"empotrado"

# Ancho fijo; altura por SizeToContent (crece al añadir capas).
_UI_WIDTH = 960
_UI_MIN_WIDTH = 920
_UI_SIDE_W = 320
_UI_CANVAS_H = 340
# Grosor trazo elevación (barras finas; muro/cortes un poco más visibles).
_STROKE_BAR = 1.25
_STROKE_WALL = 1.5
_STROKE_CUT = 1.4


def _ui_min_window_height():
    """Piso de MinHeight (~1 capa); la altura real la mide SizeToContent."""
    # Cinta SO ~32; pad 18×2; título ~40; panel pad 12×2; hint ~28; footer ~52
    chrome = 200
    stack = 28 + 40 + 22 + 54 + 82 + 24
    if ENABLE_TRASLAPOS:
        stack += 118
    body = max(_UI_CANVAS_H + 34, stack) + 8
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
        return u"AraincoCoronamientoMurosColocar"

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


class CoronamientoMurosWindow(object):
    def __init__(
        self,
        uiapp,
        uidoc,
        doc,
        wall,
        on_closed=None,
        geom_mode=None,
        upper_wall=None,
        walls_ord=None,
        voladizo_specs=None,
        overhang_mm=0.0,
        embed_side=None,
    ):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        self._wall = wall
        self._on_closed = on_closed
        mode = _as_unicode(geom_mode or _GEOM_U_LIBRE).strip().lower()
        self._geom_mode = (
            _GEOM_EMPOTRADO if mode == _GEOM_EMPOTRADO else _GEOM_U_LIBRE
        )
        self._upper_wall = upper_wall
        self._walls_ord = list(walls_ord or ([wall] if wall is not None else []))
        self._voladizo_specs = list(voladizo_specs or [])
        try:
            self._overhang_mm = float(overhang_mm or 0.0)
        except Exception:
            self._overhang_mm = 0.0
        self._embed_side = embed_side
        self._layers = sync_layers(None, 1)
        self._n_capas = 1
        self._lap_mode = LAP_MODE_SYMMETRIC
        self._cuts_mm = []
        self._cmb_capas = None
        self._cmb_lap = None
        self._layer_host = None
        self._elev_draw = {}
        self._canvas = None
        self._txt_status = None
        self._txt_warn = None
        self._warn_border = None
        self._layer_combos = []  # [(cmb_n, cmb_d), ...]
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
        if ENABLE_TRASLAPOS:
            hint = u"Pick empalme en elevación · 12 m solo aviso · traslape global"
        else:
            hint = u"Capas n/Ø · 12 m solo aviso"
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

    def _is_empotrado(self):
        return self._geom_mode == _GEOM_EMPOTRADO

    def _length_estimate(self):
        if self._is_empotrado():
            diam = 16
            if self._layers:
                diam = self._layers[0].get(u"diam_mm", 16)
            return wall_length_estimate_empotrado(
                self._wall, self._overhang_mm, diam_mm=diam
            )
        return wall_length_estimate(self._wall)
    def _build_body_xaml(self):
        canvas_cursor = u"Cross" if ENABLE_TRASLAPOS else u"Arrow"
        if ENABLE_TRASLAPOS:
            if self._is_empotrado():
                stack_hint = (
                    u"Host Z baja · empotro V3 · n/Ø por capa · pick empalme"
                )
            else:
                stack_hint = (
                    u"Capas con n y Ø · empalme por clic en elevación · traslape global"
                )
            empalme_xaml = u"""
        <Border Height="1" Background="#21465C" Margin="0,8,0,8"/>
        <TextBlock Text="EMPALME (conjunto)" Foreground="#64748b" FontSize="10"
                   FontWeight="SemiBold" Margin="0,0,0,4"/>
        <ComboBox x:Name="CmbLap" Style="{StaticResource ComboStretch}"
                  HorizontalAlignment="Stretch" Margin="0,0,0,6"/>
        <TextBlock x:Name="TxtCutsHint" Foreground="#64748b" FontSize="9"
                   TextWrapping="Wrap" Margin="0,0,0,6"
                   Text="Clic en elevación: añade corte · clic cerca: quita"/>
"""
        else:
            if self._is_empotrado():
                stack_hint = u"Host Z baja · empotro V3 · capas n/Ø"
            else:
                stack_hint = u"Capas con n y Ø"
            empalme_xaml = u""
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
        <RowDefinition Height="__CANVAS_H__"/>
      </Grid.RowDefinitions>
      <DockPanel Grid.Row="0" Margin="8,6">
        <TextBlock x:Name="TxtElevHdr" DockPanel.Dock="Left" Text="Elevación"
                   Foreground="#95B8CC" FontSize="11" VerticalAlignment="Center"/>
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
          <Border x:Name="ChipMode" BorderBrush="#fbbf24" BorderThickness="1"
                  Background="#0E1B32" Padding="6,2" CornerRadius="3" Margin="0,0,6,0">
            <TextBlock x:Name="TxtModeChip" Text="U LIBRE" Foreground="#fbbf24"
                       FontSize="10" FontWeight="SemiBold"/>
          </Border>
          <Border BorderBrush="#fbbf24" BorderThickness="1"
                  Background="#0E1B32" Padding="6,2" CornerRadius="3">
            <TextBlock Text="COR" Foreground="#fbbf24" FontSize="10" FontWeight="SemiBold"/>
          </Border>
        </StackPanel>
      </DockPanel>
      <Canvas x:Name="CnvElev" Grid.Row="1" Background="#050E18" ClipToBounds="True"
              Height="__CANVAS_H__" MinHeight="__CANVAS_H__" Cursor="__CURSOR__"/>
    </Grid>
  </Border>
  <Border x:Name="PanelStack" Grid.Column="1" Background="#0a1620" BorderBrush="#fbbf24" BorderThickness="1.5"
          CornerRadius="4" Padding="10" VerticalAlignment="Top">
    <StackPanel>
        <DockPanel Margin="0,0,0,6">
          <TextBlock Text="Stack coronamiento" Foreground="#E8F4F8" FontSize="12"
                     FontWeight="SemiBold" VerticalAlignment="Center"/>
          <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
            <Border x:Name="ChipModeSide" BorderBrush="#fbbf24" BorderThickness="1"
                    Background="#0E1B32" Padding="6,2" CornerRadius="3" Margin="0,0,6,0">
              <TextBlock x:Name="TxtModeChipSide" Text="U LIBRE" Foreground="#fbbf24"
                         FontSize="10" FontWeight="SemiBold"/>
            </Border>
            <Border BorderBrush="#fbbf24" BorderThickness="1"
                    Background="#0E1B32" Padding="6,2" CornerRadius="3">
              <TextBlock Text="COR" Foreground="#fbbf24" FontSize="10" FontWeight="SemiBold"/>
            </Border>
          </StackPanel>
        </DockPanel>
        <TextBlock Text="__HINT__"
                   Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,0,8"/>
        <TextBlock x:Name="TxtWallChip" Foreground="#95B8CC" FontSize="10"
                   Margin="0,0,0,8" TextWrapping="Wrap"/>
        <TextBlock Text="CAPAS" Foreground="#64748b" FontSize="10" FontWeight="SemiBold"
                   Margin="0,0,0,4"/>
        <ComboBox x:Name="CmbCapas" Style="{StaticResource ComboStretch}"
                  HorizontalAlignment="Stretch" Margin="0,0,0,8"/>
        <StackPanel x:Name="PanelLayers"/>
__EMPALME__
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
        xaml = xaml.replace(u"__CURSOR__", canvas_cursor)
        xaml = xaml.replace(u"__HINT__", stack_hint)
        xaml = xaml.replace(u"__EMPALME__", empalme_xaml)
        return xaml.strip()

    def _fit_height_to_content(self):
        """SizeToContent=Height → medir → Manual con Height real (sin forzar 3 capas)."""
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
        """Tras el primer layout, fija altura al contenido real (sin aire extra)."""
        try:
            self._fit_height_to_content()
            self._redraw()
        except Exception:
            pass

    def _wire(self):
        w = self._win
        self._canvas = w.FindName(u"CnvElev")
        self._cmb_capas = w.FindName(u"CmbCapas")
        self._cmb_lap = w.FindName(u"CmbLap") if ENABLE_TRASLAPOS else None
        self._layer_host = w.FindName(u"PanelLayers")
        self._txt_status = w.FindName(u"TxtStatus")
        self._txt_warn = w.FindName(u"TxtWarn")
        self._warn_border = w.FindName(u"WarnBorder")
        chip = w.FindName(u"TxtWallChip")
        try:
            wid = self._wall.Id.IntegerValue
        except Exception:
            wid = u"?"
        if chip is not None:
            if self._is_empotrado() and self._upper_wall is not None:
                try:
                    uid = self._upper_wall.Id.IntegerValue
                except Exception:
                    uid = u"?"
                chip.Text = u"Host Id {0} · Sup Id {1} · EMPOTRADO".format(wid, uid)
            else:
                chip.Text = u"Muro Id {0} · U LIBRE".format(wid)
        self._apply_mode_chips()

        for n in CAPAS_OPTS:
            it = ComboBoxItem()
            it.Content = u"{0}".format(n)
            it.Tag = int(n)
            self._cmb_capas.Items.Add(it)
        self._cmb_capas.SelectedIndex = 0

        if ENABLE_TRASLAPOS and self._cmb_lap is not None:
            for key, label in LAP_MODE_LABELS:
                it = ComboBoxItem()
                it.Content = label
                it.Tag = key
                self._cmb_lap.Items.Add(it)
            self._cmb_lap.SelectedIndex = 0

        def on_capas(sender, args):
            it = self._cmb_capas.SelectedItem
            if it is None:
                return
            n = clamp_n_capas(it.Tag)
            self._n_capas = n
            self._layers = sync_layers(self._layers, n)
            self._rebuild_layer_cards()
            self._redraw()
            self._set_status(self._status_line())
            self._fit_height_to_content()

        def on_lap(sender, args):
            it = self._cmb_lap.SelectedItem
            if it is None:
                return
            self._lap_mode = normalize_lap_mode_ui(it.Tag)
            self._redraw()
            self._set_status(self._status_line())

        from System.Windows.Controls import SelectionChangedEventHandler

        self._cmb_capas.SelectionChanged += SelectionChangedEventHandler(on_capas)
        if ENABLE_TRASLAPOS and self._cmb_lap is not None:
            self._cmb_lap.SelectionChanged += SelectionChangedEventHandler(on_lap)

        if self._canvas is not None:
            def _on_size(sender, args):
                self._redraw()

            self._canvas.SizeChanged += SizeChangedEventHandler(_on_size)
            if ENABLE_TRASLAPOS:
                self._canvas.MouseLeftButtonDown += MouseButtonEventHandler(
                    self._on_canvas_click
                )

        btn_c = w.FindName(u"BtnCancelar")
        btn_p = w.FindName(u"BtnColocar")
        if btn_c is not None:
            btn_c.Click += RoutedEventHandler(lambda s, a: self._win.Close())
        if btn_p is not None:
            btn_p.Click += RoutedEventHandler(lambda s, a: self._colocar_event.Raise())

    def _apply_mode_chips(self):
        is_emb = self._is_empotrado()
        label = u"EMPOTRADO" if is_emb else u"U LIBRE"
        color = _EMBED if is_emb else _COR
        for name in (u"TxtModeChip", u"TxtModeChipSide"):
            tb = self._win.FindName(name) if self._win is not None else None
            if tb is None:
                continue
            tb.Text = label
            try:
                tb.Foreground = _brush(color)
            except Exception:
                pass
        for name in (u"ChipMode", u"ChipModeSide", u"PanelStack"):
            br = self._win.FindName(name) if self._win is not None else None
            if br is None:
                continue
            try:
                br.BorderBrush = _brush(color)
            except Exception:
                pass

    def _rebuild_layer_cards(self):
        host = self._layer_host
        if host is None:
            return
        host.Children.Clear()
        self._layer_combos = []
        from System.Windows.Controls import SelectionChangedEventHandler

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
            hdr.Text = u"{0}ª C.".format(i + 1)
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

            def _mk_combo(values, selected, width=70):
                cmb = ComboBox()
                try:
                    cmb.Style = self._win.FindResource(u"Combo")
                except Exception:
                    pass
                cmb.Width = width
                cmb.Margin = Thickness(0, 0, 8, 0)
                for v in values:
                    it = ComboBoxItem()
                    it.Content = u"{0}".format(v) if width < 80 else u"Ø{0}".format(v)
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
            cmb_d = _mk_combo(DIAMS_MM, ly.get(u"diam_mm", 16), 72)

            li = i

            def _on_change(sender, args, idx=li):
                self._read_layer_combos()
                self._redraw()
                self._refresh_warn()
                self._set_status(self._status_line())

            cmb_n.SelectionChanged += SelectionChangedEventHandler(_on_change)
            cmb_d.SelectionChanged += SelectionChangedEventHandler(_on_change)

            row.Children.Add(lbl_n)
            row.Children.Add(cmb_n)
            row.Children.Add(lbl_d)
            row.Children.Add(cmb_d)
            sp.Children.Add(row)

            tip = TextBlock()
            tip.Text = u"Traslape Ø ≈ {0} mm".format(
                format_mm_es(traslape_mm_from_diam(ly.get(u"diam_mm", 16)))
            )
            tip.Foreground = _brush(u"#64748b")
            tip.FontSize = 9
            tip.Margin = Thickness(0, 4, 0, 0)
            tip.Tag = u"tip"
            if not ENABLE_TRASLAPOS:
                tip.Visibility = Visibility.Collapsed
            sp.Children.Add(tip)

            card.Child = sp
            host.Children.Add(card)
            self._layer_combos.append((cmb_n, cmb_d, tip))

    def _read_layer_combos(self):
        layers = []
        for cmb_n, cmb_d, tip in self._layer_combos:
            n = 2
            d = 16
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
            if tip is not None:
                tip.Text = u"Traslape Ø ≈ {0} mm".format(
                    format_mm_es(traslape_mm_from_diam(d))
                )
        if layers:
            self._layers = layers

    def _status_line(self):
        mode = u"Empotrado" if self._is_empotrado() else u"U libre"
        parts = []
        for i, ly in enumerate(self._layers):
            parts.append(
                u"C{0}:{1}Ø{2}".format(i + 1, ly.get(u"n_bars"), ly.get(u"diam_mm"))
            )
        base = u"{0} · {1}".format(mode, u" · ".join(parts))
        if not ENABLE_TRASLAPOS:
            return base
        ncuts = len(self._cuts_mm)
        return u"{0} · {1} corte(s) · {2}".format(base, ncuts, self._lap_mode)

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _refresh_warn(self):
        est = self._length_estimate()
        if self._warn_border is None or self._txt_warn is None:
            return
        if est.get(u"exceeds_12m"):
            self._warn_border.Visibility = Visibility.Visible
            if self._is_empotrado():
                self._txt_warn.Text = (
                    u">12 m · L≈{0} mm (aviso). "
                    u"Empotro {1} + voladizo {2} + 1×pata {3} mm."
                ).format(
                    format_mm_es(est[u"developed_mm"]),
                    format_mm_es(est.get(u"embed_mm", 0)),
                    format_mm_es(est.get(u"overhang_mm", 0)),
                    format_mm_es(est[u"pata_mm"]),
                )
            elif ENABLE_TRASLAPOS:
                self._txt_warn.Text = (
                    u">12 m · L≈{0} mm (aviso; no bloquea traslape). "
                    u"Vano {1} mm + 2×pata {2} mm."
                ).format(
                    format_mm_es(est[u"developed_mm"]),
                    format_mm_es(est[u"main_mm"]),
                    format_mm_es(est[u"pata_mm"]),
                )
            else:
                self._txt_warn.Text = (
                    u">12 m · L≈{0} mm (aviso). "
                    u"Vano {1} mm + 2×pata {2} mm."
                ).format(
                    format_mm_es(est[u"developed_mm"]),
                    format_mm_es(est[u"main_mm"]),
                    format_mm_es(est[u"pata_mm"]),
                )
        else:
            self._warn_border.Visibility = Visibility.Collapsed
            if self._is_empotrado():
                self._txt_warn.Text = (
                    u"L≈{0} mm · empotro+voladizo+1 pata."
                ).format(format_mm_es(est[u"developed_mm"]))
            elif ENABLE_TRASLAPOS:
                self._txt_warn.Text = (
                    u"L≈{0} mm · clic en elevación para empalme."
                ).format(format_mm_es(est[u"developed_mm"]))
            else:
                self._txt_warn.Text = u"L≈{0} mm.".format(
                    format_mm_es(est[u"developed_mm"])
                )

    def _active_view(self):
        try:
            if self._uidoc is not None:
                return self._uidoc.ActiveView
        except Exception:
            pass
        return None

    def _canvas_flip(self):
        """True si izq. del canvas = extremo P1 en pantalla (mm=main a la izq.)."""
        return bool(
            wall_elev_canvas_flip_for_view(self._wall, self._active_view())
        )

    def _on_canvas_click(self, sender, args):
        if not ENABLE_TRASLAPOS:
            return
        if self._canvas is None:
            return
        est = self._length_estimate()
        main = float(est[u"main_mm"])
        if main <= 1.0:
            return
        pt = args.GetPosition(self._canvas)
        w = float(self._canvas.ActualWidth or 0)
        h = float(self._canvas.ActualHeight or 0)
        if w < 10 or h < 10:
            return
        draw = getattr(self, u"_elev_draw", None) or {}
        bar_x0 = draw.get(u"bar_x0")
        bar_x1 = draw.get(u"bar_x1")
        if bar_x0 is None or bar_x1 is None:
            # Empotrado: barra no ocupa todo el muro — mismos insets que redraw.
            if self._is_empotrado():
                bar_x0, bar_x1 = self._empotrado_bar_x_range(w, main, est)
            else:
                pad_l, pad_r = 48.0, 28.0
                bar_x0 = pad_l + 18.0
                bar_x1 = w - pad_r - 18.0
        span = max(1.0, float(bar_x1) - float(bar_x0))
        x = float(pt.X)
        if x < float(bar_x0) or x > float(bar_x1):
            return
        t = (x - float(bar_x0)) / span
        if self._canvas_flip():
            mm = (1.0 - t) * main
        else:
            mm = t * main
        self._cuts_mm = toggle_cut_at_mm(self._cuts_mm, mm, main)
        self._redraw()
        self._set_status(self._status_line())

    def _elev_walls_for_layout(self):
        """Muros en orden base→cima para proyección (host, [upper])."""
        if self._is_empotrado() and self._upper_wall is not None:
            if self._walls_ord and len(self._walls_ord) >= 2:
                return list(self._walls_ord[:2])
            return [self._wall, self._upper_wall]
        if self._wall is not None:
            return [self._wall]
        return []

    def _elev_layout_and_rects(self, canvas_w, canvas_h):
        """
        Layout modelo → rectángulos canvas (px).

        X = proyección Right (menor u → izq.); Y = Z (mayor z → arriba).
        Escala uniforme px/ft (misma para X y Z) para conservar proporciones
        y offsets relativos entre muros; el bloque se centra en la zona útil.
        """
        pad_l, pad_r, pad_t, pad_b = 48.0, 28.0, 28.0, 36.0
        wx0, wx1 = pad_l, float(canvas_w) - pad_r
        wy0, wy1 = pad_t, float(canvas_h) - pad_b
        usable_w = max(20.0, wx1 - wx0)
        usable_h = max(20.0, wy1 - wy0)
        walls = self._elev_walls_for_layout()
        layout = walls_elev_layout_model(walls, self._active_view())
        out = {
            u"pad": (pad_l, pad_r, pad_t, pad_b),
            u"zone": (wx0, wx1, wy0, wy1),
            u"layout": layout,
            u"rects": [],
            u"host_rect": None,
            u"upper_rect": None,
        }
        if layout is None or not layout.get(u"items"):
            host_r = {
                u"x0": wx0,
                u"x1": wx1,
                u"y0": wy0,
                u"y1": wy1,
                u"item": None,
                u"wall": self._wall,
                u"is_host": True,
            }
            out[u"rects"] = [host_r]
            out[u"host_rect"] = host_r
            return out

        g_u_min = float(layout[u"global_u_min"])
        span_u = float(layout[u"global_span_u"])
        g_z_min = float(layout[u"global_z_min"])
        g_z_max = float(layout[u"global_z_max"])
        span_z = float(layout[u"global_span_z"])

        # Escala uniforme: el stack/muro cabe en la zona sin deformar.
        sx = usable_w / span_u
        sz = usable_h / span_z
        scale = min(sx, sz)
        if scale < 1e-12:
            scale = 1e-12
        draw_w = span_u * scale
        draw_h = span_z * scale
        ox = wx0 + 0.5 * (usable_w - draw_w)
        oy = wy0 + 0.5 * (usable_h - draw_h)

        def x_of(u):
            return ox + (float(u) - g_u_min) * scale

        def y_of(z):
            # Canvas Y crece hacia abajo; z_top → y menor.
            return oy + (g_z_max - float(z)) * scale

        rects = []
        for i, item in enumerate(layout[u"items"]):
            x0 = x_of(item[u"u_start"])
            x1 = x_of(item[u"u_end"])
            y_top = y_of(item[u"z_top"])
            y_bot = y_of(item[u"z_bot"])
            if x1 < x0:
                x0, x1 = x1, x0
            if y_bot < y_top:
                y_top, y_bot = y_bot, y_top
            if x1 - x0 < 4.0:
                mid = 0.5 * (x0 + x1)
                x0, x1 = mid - 2.0, mid + 2.0
            if y_bot - y_top < 4.0:
                mid = 0.5 * (y_top + y_bot)
                y_top, y_bot = mid - 2.0, mid + 2.0
            is_host = i == 0
            r = {
                u"x0": float(x0),
                u"x1": float(x1),
                u"y0": float(y_top),
                u"y1": float(y_bot),
                u"item": item,
                u"wall": item.get(u"wall") or (walls[i] if i < len(walls) else None),
                u"is_host": is_host,
            }
            rects.append(r)
            if is_host:
                out[u"host_rect"] = r
            else:
                out[u"upper_rect"] = r
        out[u"rects"] = rects
        if out[u"host_rect"] is None and rects:
            out[u"host_rect"] = rects[0]
            rects[0][u"is_host"] = True
        return out

    def _empotrado_bar_x_range(self, canvas_w, main_mm, est, host_rect=None, upper_rect=None):
        """Rango X de la barra L en canvas (empotro+voladizo) sobre host proyectado."""
        pad_l, pad_r = 48.0, 28.0
        end_inset = 18.0
        embed = float(est.get(u"embed_mm") or 0.0)
        overhang = float(est.get(u"overhang_mm") or self._overhang_mm or 0.0)
        Lhost = max(1.0, wall_largo_mm(self._wall))

        if host_rect is None:
            # Sin rectos: estimar desde layout o zona completa.
            try:
                info = self._elev_layout_and_rects(canvas_w, 200.0)
                host_rect = info.get(u"host_rect")
                upper_rect = info.get(u"upper_rect")
            except Exception:
                host_rect = None
                upper_rect = None

        if host_rect is not None:
            hx0 = float(host_rect[u"x0"])
            hx1 = float(host_rect[u"x1"])
        else:
            hx0 = pad_l
            hx1 = float(canvas_w) - pad_r
        wall_w = max(1.0, hx1 - hx0)

        # Lado de voladizo: geometría superior vs host; respaldo embed_side.
        side = self._embed_side or u"der"
        if upper_rect is not None:
            ux0 = float(upper_rect[u"x0"])
            ux1 = float(upper_rect[u"x1"])
            oh_right = max(0.0, hx1 - ux1)
            oh_left = max(0.0, ux0 - hx0)
            if oh_left > oh_right + 1.0:
                side = u"izq"
            elif oh_right > oh_left + 1.0:
                side = u"der"

        if side == u"izq":
            # Voladizo a la izquierda: barra desde extremo libre hacia reentrada.
            bar_x0 = hx0 + end_inset
            bar_x1 = bar_x0 + max(
                20.0, (float(main_mm) / max(1.0, Lhost)) * wall_w
            )
            if upper_rect is not None:
                reent_x = float(upper_rect[u"x0"])
                bar_x1 = max(
                    bar_x0 + 20.0,
                    min(hx1 - end_inset, reent_x + (embed / max(1.0, Lhost)) * wall_w),
                )
            else:
                bar_x1 = min(hx1 - end_inset, bar_x1)
        else:
            # Empotro bajo superior (izq) + voladizo libre (der).
            if upper_rect is not None:
                reent_x = float(upper_rect[u"x1"])
            else:
                reent_frac = max(
                    0.05, min(0.95, (Lhost - overhang) / max(1.0, Lhost))
                )
                reent_x = hx0 + wall_w * reent_frac
            bar_x0 = max(hx0 + 8.0, reent_x - (embed / max(1.0, Lhost)) * wall_w)
            bar_x1 = hx1 - end_inset
        return float(bar_x0), float(bar_x1)

    def _clear_canvas(self):
        if self._canvas is not None:
            self._canvas.Children.Clear()

    def _add_line(self, x1, y1, x2, y2, color, thick=1.5, dash=None):
        ln = Line()
        ln.X1, ln.Y1, ln.X2, ln.Y2 = x1, y1, x2, y2
        ln.Stroke = _brush(color)
        ln.StrokeThickness = thick
        if dash:
            from System.Windows.Media import DoubleCollection

            dc = DoubleCollection()
            for v in dash:
                dc.Add(float(v))
            ln.StrokeDashArray = dc
        self._canvas.Children.Add(ln)

    def _add_rect(self, x, y, w, h, fill, stroke=None, stroke_w=1.0):
        r = Rectangle()
        r.Width = max(0.5, w)
        r.Height = max(0.5, h)
        Canvas.SetLeft(r, x)
        Canvas.SetTop(r, y)
        r.Fill = fill
        if stroke:
            r.Stroke = _brush(stroke)
            r.StrokeThickness = stroke_w
        self._canvas.Children.Add(r)

    def _add_text(self, x, y, text, color, size=10):
        tb = TextBlock()
        tb.Text = _as_unicode(text)
        tb.Foreground = _brush(color)
        tb.FontSize = size
        Canvas.SetLeft(tb, x)
        Canvas.SetTop(tb, y)
        self._canvas.Children.Add(tb)

    def _redraw(self):
        if self._canvas is None:
            return
        self._clear_canvas()
        w = float(self._canvas.ActualWidth or 0)
        h = float(self._canvas.ActualHeight or 0)
        if w < 40 or h < 40:
            self._elev_draw = {}
            return
        est = self._length_estimate()
        main = float(est[u"main_mm"])
        is_emb = self._is_empotrado()
        info = self._elev_layout_and_rects(w, h)
        pad_l, pad_r, pad_t, pad_b = info[u"pad"]
        wx0, wx1, wy0, wy1 = info[u"zone"]
        host_r = info.get(u"host_rect")
        upper_r = info.get(u"upper_rect")
        if host_r is None:
            self._elev_draw = {}
            return

        hx0, hx1 = float(host_r[u"x0"]), float(host_r[u"x1"])
        hy0, hy1 = float(host_r[u"y0"]), float(host_r[u"y1"])
        host_top_y = hy0  # tope del host (barras aquí)

        # Host (inferior / único)
        self._add_rect(
            hx0,
            hy0,
            max(4.0, hx1 - hx0),
            max(4.0, hy1 - hy0),
            _brush(_WALL),
            _WALL_STROKE,
            _STROKE_WALL,
        )

        if is_emb and upper_r is not None:
            ux0, ux1 = float(upper_r[u"x0"]), float(upper_r[u"x1"])
            uy0, uy1 = float(upper_r[u"y0"]), float(upper_r[u"y1"])
            self._add_rect(
                ux0,
                uy0,
                max(4.0, ux1 - ux0),
                max(4.0, uy1 - uy0),
                _brush(_UPPER_FILL),
                _UPPER_STROKE,
                _STROKE_WALL,
            )
            self._add_text(
                (ux0 + ux1) * 0.5 - 36,
                uy0 + 4,
                u"superior",
                _UPPER_STROKE,
                8,
            )
            # Contacto stack (interfaz host top ≈ upper base)
            iface_y = hy0
            self._add_line(
                min(hx0, ux0) - 8,
                iface_y,
                max(hx1, ux1) + 4,
                iface_y,
                _EMBED,
                1.0,
                (3, 2),
            )

        n_capas = len(self._layers)
        gap = 10.0
        cover = 10.0
        thick = float(_STROKE_BAR)
        leg = max(1.0, thick * 0.9)
        flip = self._canvas_flip()
        embed_mm = float(est.get(u"embed_mm") or 0.0)

        if is_emb:
            bar_x0, bar_x1 = self._empotrado_bar_x_range(
                w, main, est, host_rect=host_r, upper_rect=upper_r
            )
        else:
            bar_x0 = hx0 + 18.0
            bar_x1 = hx1 - 18.0
            if bar_x1 <= bar_x0 + 4.0:
                bar_x0 = hx0 + 4.0
                bar_x1 = hx1 - 4.0
        span = max(1.0, bar_x1 - bar_x0)
        self._elev_draw = {
            u"bar_x0": float(bar_x0),
            u"bar_x1": float(bar_x1),
            u"host_rect": host_r,
            u"upper_rect": upper_r,
            u"host_top_y": float(host_top_y),
            u"flip": bool(flip),
        }

        def x_at(mm):
            t = float(mm) / max(1.0, main)
            if flip:
                t = 1.0 - t
            return bar_x0 + t * span

        # Zona empotrada (solo Empotrado)
        if is_emb and embed_mm > 1.0 and main > 1.0:
            embed_end = min(main, embed_mm)
            xa, xb = x_at(0.0), x_at(embed_end)
            x_emb = min(xa, xb)
            self._add_rect(
                x_emb,
                host_top_y,
                max(2.0, abs(xb - xa)),
                cover + max(0, n_capas - 1) * gap + 16,
                _brush(_EMBED, 40),
                _EMBED,
                0.8,
            )
            self._add_text(
                x_emb + 2,
                host_top_y - 14 if host_top_y > 20 else host_top_y + 2,
                u"empotro",
                _EMBED,
                8,
            )

        for li, ly in enumerate(self._layers):
            y = host_top_y + cover + li * gap
            color = _LAYER_COLORS[li % len(_LAYER_COLORS)]
            diam = float(ly.get(u"diam_mm", 16))
            lap = traslape_mm_from_diam(diam)
            cuts = (
                stagger_cuts_for_layer(self._cuts_mm, li, main, lap)
                if ENABLE_TRASLAPOS
                else []
            )
            for c in cuts:
                if self._lap_mode == u"endpoint_prev":
                    a, b = c, c + lap
                elif self._lap_mode == u"endpoint_next":
                    a, b = c - lap, c
                else:
                    a, b = c - lap * 0.5, c + lap * 0.5
                a = max(0.0, a)
                b = min(main, b)
                xa, xb = x_at(a), x_at(b)
                x_lap = min(xa, xb)
                self._add_rect(
                    x_lap,
                    y - 4,
                    max(2.0, abs(xb - xa)),
                    thick + 8,
                    _brush(color, 60),
                )
            x_start = x_at(0.0)
            x_end = x_at(main) if main > 1.0 else x_at(1.0)
            self._add_line(x_start, y, x_end, y, color, thick)
            # U libre: 2 patas; Empotrado: solo pata en extremo libre (fin mm)
            if is_emb:
                el = Ellipse()
                el.Width = 7
                el.Height = 7
                el.Stroke = _brush(_EMBED)
                el.StrokeThickness = 1.4
                el.Fill = _brush(u"#050E18")
                Canvas.SetLeft(el, x_start - 3.5)
                Canvas.SetTop(el, y - 3.5)
                self._canvas.Children.Add(el)
                self._add_line(x_end, y, x_end, y + 18, color, leg)
            else:
                self._add_line(x_start, y, x_start, y + 18, color, leg)
                self._add_line(x_end, y, x_end, y + 18, color, leg)
            for ci, c in enumerate(cuts):
                self._add_line(
                    x_at(c), host_top_y - 2, x_at(c), y + 14, _CUT, _STROKE_CUT, (3, 2)
                )
                if li == 0:
                    self._add_text(
                        x_at(c) - 8, host_top_y - 16, u"C{0}".format(ci + 1), _CUT, 9
                    )

        foot = (
            u"L≈{0} mm (empotro+voladizo+1 pata)".format(
                format_mm_es(est[u"developed_mm"])
            )
            if is_emb
            else u"{0} mm · L≈{1} mm".format(
                format_mm_es(main), format_mm_es(est[u"developed_mm"])
            )
        )
        self._add_text(
            wx0 + (wx1 - wx0) * 0.5 - 60,
            wy1 + 8,
            foot,
            u"#64748b",
            9,
        )
        hdr = self._win.FindName(u"TxtElevHdr")
        if hdr is not None:
            mode = u"Empotrado" if is_emb else u"U libre"
            if ENABLE_TRASLAPOS:
                hdr.Text = u"Elevación · {0} · {1} capa(s) · {2} corte(s)".format(
                    mode, n_capas, len(self._cuts_mm)
                )
            else:
                hdr.Text = u"Elevación · {0} · {1} capa(s)".format(mode, n_capas)

    def show(self):
        self._win.Show()

    def set_wall(self, wall):
        self._wall = wall
        self._cuts_mm = []
        chip = self._win.FindName(u"TxtWallChip")
        try:
            wid = wall.Id.IntegerValue
        except Exception:
            wid = u"?"
        if chip is not None:
            if self._is_empotrado() and self._upper_wall is not None:
                try:
                    uid = self._upper_wall.Id.IntegerValue
                except Exception:
                    uid = u"?"
                chip.Text = u"Host Id {0} · Sup Id {1} · EMPOTRADO".format(wid, uid)
            else:
                chip.Text = u"Muro Id {0} · U LIBRE".format(wid)
        self._refresh_warn()
        self._redraw()
        self._set_status(self._status_line())

    def _execute_colocar(self):
        self._read_layer_combos()
        cuts = list(self._cuts_mm) if ENABLE_TRASLAPOS else []
        res = place_coronamiento(
            self._doc,
            self._uidoc,
            self._wall,
            self._layers,
            cuts_ref_mm=cuts,
            lap_mode_ui=self._lap_mode if ENABLE_TRASLAPOS else None,
            geom_mode=self._geom_mode,
            voladizo_specs=self._voladizo_specs,
            overhang_mm=self._overhang_mm,
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


def show_coronamiento_window(
    uiapp,
    uidoc,
    doc,
    wall,
    on_closed=None,
    geom_mode=None,
    upper_wall=None,
    walls_ord=None,
    voladizo_specs=None,
    overhang_mm=0.0,
    embed_side=None,
):
    if _focus_existing():
        return None
    win = CoronamientoMurosWindow(
        uiapp,
        uidoc,
        doc,
        wall,
        on_closed=on_closed,
        geom_mode=geom_mode,
        upper_wall=upper_wall,
        walls_ord=walls_ord,
        voladizo_specs=voladizo_specs,
        overhang_mm=overhang_mm,
        embed_side=embed_side,
    )
    win.show()
    return win
