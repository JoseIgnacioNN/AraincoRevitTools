# -*- coding: utf-8 -*-
"""
UI WPF — Arainco: Coronamiento muros.

Elevación + stack lateral (capas n/Ø estilo V3), pick de empalme en canvas,
escenario de traslape global (mismos modos que 56_DividirRebarPuntoTraslape),
aviso >12 m (no bloquea).

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
from System.Windows.Shapes import Line, Rectangle
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
    place_coronamiento_wall,
    wall_elev_canvas_flip_for_view,
    wall_length_estimate,
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
_LAYER_COLORS = (u"#fbbf24", u"#38bdf8", u"#a78bfa")

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
    def __init__(self, uiapp, uidoc, doc, wall, on_closed=None):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        self._wall = wall
        self._on_closed = on_closed
        self._layers = sync_layers(None, 1)
        self._n_capas = 1
        self._lap_mode = LAP_MODE_SYMMETRIC
        self._cuts_mm = []
        self._cmb_capas = None
        self._cmb_lap = None
        self._layer_host = None
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

    def _build_body_xaml(self):
        canvas_cursor = u"Cross" if ENABLE_TRASLAPOS else u"Arrow"
        if ENABLE_TRASLAPOS:
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
        <Border DockPanel.Dock="Right" BorderBrush="#fbbf24" BorderThickness="1"
                Background="#0E1B32" Padding="6,2" CornerRadius="3">
          <TextBlock Text="COR" Foreground="#fbbf24" FontSize="10" FontWeight="SemiBold"/>
        </Border>
      </DockPanel>
      <Canvas x:Name="CnvElev" Grid.Row="1" Background="#050E18" ClipToBounds="True"
              Height="__CANVAS_H__" MinHeight="__CANVAS_H__" Cursor="__CURSOR__"/>
    </Grid>
  </Border>
  <Border Grid.Column="1" Background="#0a1620" BorderBrush="#fbbf24" BorderThickness="1.5"
          CornerRadius="4" Padding="10" VerticalAlignment="Top">
    <StackPanel>
        <DockPanel Margin="0,0,0,6">
          <TextBlock Text="Stack coronamiento" Foreground="#E8F4F8" FontSize="12"
                     FontWeight="SemiBold" VerticalAlignment="Center"/>
          <Border DockPanel.Dock="Right" BorderBrush="#fbbf24" BorderThickness="1"
                  Background="#0E1B32" Padding="6,2" CornerRadius="3">
            <TextBlock Text="COR" Foreground="#fbbf24" FontSize="10" FontWeight="SemiBold"/>
          </Border>
        </DockPanel>
        <TextBlock Text="__HINT__"
                   Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,0,8"/>
        <TextBlock x:Name="TxtWallChip" Foreground="#95B8CC" FontSize="10"
                   Margin="0,0,0,8"/>
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
            chip.Text = u"Muro Id {0}".format(wid)

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
        parts = []
        for i, ly in enumerate(self._layers):
            parts.append(
                u"C{0}:{1}Ø{2}".format(i + 1, ly.get(u"n_bars"), ly.get(u"diam_mm"))
            )
        if not ENABLE_TRASLAPOS:
            return u" · ".join(parts)
        ncuts = len(self._cuts_mm)
        return u"{0} · {1} corte(s) · {2}".format(
            u" · ".join(parts), ncuts, self._lap_mode
        )

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _refresh_warn(self):
        est = wall_length_estimate(self._wall)
        if self._warn_border is None or self._txt_warn is None:
            return
        if est[u"exceeds_12m"]:
            self._warn_border.Visibility = Visibility.Visible
            if ENABLE_TRASLAPOS:
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
            if ENABLE_TRASLAPOS:
                self._txt_warn.Text = (
                    u"L≈{0} mm · clic en elevación para empalme."
                ).format(format_mm_es(est[u"developed_mm"]))
            else:
                self._txt_warn.Text = u"L≈{0} mm.".format(
                    format_mm_es(est[u"developed_mm"])
                )

    def _on_canvas_click(self, sender, args):
        if not ENABLE_TRASLAPOS:
            return
        if self._canvas is None:
            return
        est = wall_length_estimate(self._wall)
        main = float(est[u"main_mm"])
        if main <= 1.0:
            return
        pt = args.GetPosition(self._canvas)
        w = float(self._canvas.ActualWidth or 0)
        h = float(self._canvas.ActualHeight or 0)
        if w < 10 or h < 10:
            return
        pad_l, pad_r = 48.0, 28.0
        bar_x0 = pad_l + 18.0
        bar_x1 = w - pad_r - 18.0
        span = max(1.0, bar_x1 - bar_x0)
        x = float(pt.X)
        if x < bar_x0 or x > bar_x1:
            return
        mm = (x - bar_x0) / span * main
        self._cuts_mm = toggle_cut_at_mm(self._cuts_mm, mm, main)
        self._redraw()
        self._set_status(self._status_line())

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
            return
        est = wall_length_estimate(self._wall)
        main = float(est[u"main_mm"])
        pad_l, pad_r, pad_t, pad_b = 48.0, 28.0, 28.0, 36.0
        wx0, wx1 = pad_l, w - pad_r
        wy0, wy1 = pad_t, h - pad_b
        wall_w = wx1 - wx0
        wall_h = wy1 - wy0
        self._add_rect(
            wx0, wy0, wall_w, wall_h, _brush(_WALL), _WALL_STROKE, _STROKE_WALL
        )

        n_capas = len(self._layers)
        gap = 10.0
        cover = 10.0
        bar_x0 = wx0 + 18.0
        bar_x1 = wx1 - 18.0
        span = max(1.0, bar_x1 - bar_x0)
        thick = float(_STROKE_BAR)
        leg = max(1.0, thick * 0.9)

        def x_at(mm):
            return bar_x0 + (float(mm) / max(1.0, main)) * span

        for li, ly in enumerate(self._layers):
            y = wy0 + cover + li * gap
            color = _LAYER_COLORS[li % len(_LAYER_COLORS)]
            diam = float(ly.get(u"diam_mm", 16))
            lap = traslape_mm_from_diam(diam)
            cuts = (
                stagger_cuts_for_layer(self._cuts_mm, li, main, lap)
                if ENABLE_TRASLAPOS
                else []
            )
            # zona lap
            for c in cuts:
                if self._lap_mode == u"endpoint_prev":
                    a, b = c, c + lap
                elif self._lap_mode == u"endpoint_next":
                    a, b = c - lap, c
                else:
                    a, b = c - lap * 0.5, c + lap * 0.5
                a = max(0.0, a)
                b = min(main, b)
                self._add_rect(
                    x_at(a),
                    y - 4,
                    max(2.0, x_at(b) - x_at(a)),
                    thick + 8,
                    _brush(color, 60),
                )
            self._add_line(bar_x0, y, bar_x1, y, color, thick)
            # patas
            self._add_line(bar_x0, y, bar_x0, y + 18, color, leg)
            self._add_line(bar_x1, y, bar_x1, y + 18, color, leg)
            for ci, c in enumerate(cuts):
                self._add_line(
                    x_at(c), wy0 - 2, x_at(c), y + 14, _CUT, _STROKE_CUT, (3, 2)
                )
                if li == 0:
                    self._add_text(x_at(c) - 8, wy0 - 16, u"C{0}".format(ci + 1), _CUT, 9)

        self._add_text(
            wx0 + wall_w * 0.5 - 40,
            wy1 + 8,
            u"{0} mm · L≈{1} mm".format(
                format_mm_es(main), format_mm_es(est[u"developed_mm"])
            ),
            u"#64748b",
            9,
        )
        hdr = self._win.FindName(u"TxtElevHdr")
        if hdr is not None:
            if ENABLE_TRASLAPOS:
                hdr.Text = u"Elevación · {0} capa(s) · {1} corte(s) ref.".format(
                    n_capas, len(self._cuts_mm)
                )
            else:
                hdr.Text = u"Elevación · {0} capa(s)".format(n_capas)

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
            chip.Text = u"Muro Id {0}".format(wid)
        self._refresh_warn()
        self._redraw()
        self._set_status(self._status_line())

    def _execute_colocar(self):
        self._read_layer_combos()
        cuts = list(self._cuts_mm) if ENABLE_TRASLAPOS else []
        res = place_coronamiento_wall(
            self._doc,
            self._uidoc,
            self._wall,
            self._layers,
            cuts_ref_mm=cuts,
            lap_mode_ui=self._lap_mode if ENABLE_TRASLAPOS else None,
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


def show_coronamiento_window(uiapp, uidoc, doc, wall, on_closed=None):
    if _focus_existing():
        return None
    win = CoronamientoMurosWindow(uiapp, uidoc, doc, wall, on_closed=on_closed)
    win.show()
    return win
