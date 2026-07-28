# -*- coding: utf-8 -*-
"""
UI WPF — Sección detalle extremo de muro.

Shell estándar BIMTools (cinta blanca Arainco) + canvas de sección transversal.
"""

from __future__ import print_function

import weakref

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import (
    WindowState,
)
from System.Windows.Controls import Canvas
from System.Windows.Input import Key, KeyEventHandler
from System.Windows.Media import (
    SolidColorBrush,
    Color,
)
from System.Windows.Shapes import Rectangle, Line, Ellipse
from System.Windows.Controls import TextBlock
from System.Windows.Markup import XamlReader
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from bimtools_instruction_dialog import show_message_dialog
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)
from seccion_detalle_extremo_muro import (
    _DIALOG_TITLE,
    _FAR_CLIP_MM_DEF,
    _LONGITUD_DETALLE_MM_DEF,
    _TX_CREAR,
    _as_unicode,
    _view_family_type_display_name,
    ejecutar_crear_detail_extremo,
    proposed_view_name,
    wall_section_preview_data,
)

_SINGLETON_KEY = u"Arainco_SeccionDetalleExtremoMuro_UI"
_TOOL_TITLE = _DIALOG_TITLE

_BODY_XAML = u"""
<StackPanel>
  <TextBlock x:Name="TxtParent" Foreground="#95B8CC" FontSize="11"
             TextWrapping="Wrap" Margin="0,0,0,8"/>
  <TextBlock x:Name="TxtWall" Foreground="#95B8CC" FontSize="11"
             TextWrapping="Wrap" Margin="0,0,0,8"/>
  <TextBlock x:Name="TxtVft" Foreground="#E8B84A" FontSize="11"
             TextWrapping="Wrap" Margin="0,0,0,10"/>

  <Border Background="#09141e" BorderBrush="#21465C" BorderThickness="1"
          CornerRadius="4" Height="280" Margin="0,0,0,12">
    <Canvas x:Name="SectionCanvas" ClipToBounds="True"/>
  </Border>

  <TextBlock Text="Extremo a detallar (confinamiento)" Foreground="#95B8CC" FontSize="11"
             FontWeight="SemiBold" Margin="0,0,0,6"/>
  <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
    <RadioButton x:Name="RbInicio" Content="Inicio (extremo 0)"
                 GroupName="Extremo" IsChecked="True" Margin="0,0,18,0"
                 Foreground="#E6EEF5" VerticalAlignment="Center"/>
    <RadioButton x:Name="RbTermino" Content="Término (extremo 1)"
                 GroupName="Extremo"
                 Foreground="#E6EEF5" VerticalAlignment="Center"/>
  </StackPanel>

  <Grid Margin="0,0,0,4">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="12"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    <StackPanel Grid.Column="0">
      <TextBlock Text="Far clip vertical (mm)" Foreground="#95B8CC" FontSize="11" Margin="0,0,0,4"/>
      <TextBox x:Name="TxtFarClip" Style="{StaticResource BimToolsTextBoxDark}"
               Text="400" Height="30"/>
    </StackPanel>
    <StackPanel Grid.Column="2">
      <TextBlock Text="Longitud en planta (mm)" Foreground="#95B8CC" FontSize="11" Margin="0,0,0,4"/>
      <TextBox x:Name="TxtLongitud" Style="{StaticResource BimToolsTextBoxDark}"
               Text="1000" Height="30"/>
    </StackPanel>
  </Grid>
  <TextBlock Text="Nombre propuesto" Foreground="#95B8CC" FontSize="11" Margin="0,8,0,4"/>
  <TextBox x:Name="TxtNombre" Style="{StaticResource BimToolsTextBoxDark}"
           IsReadOnly="True" Height="30"/>
</StackPanel>
"""

_FOOTER_ACTIONS = u"""
<Button x:Name="BtnCancel" Content="Cancelar" Margin="0,0,8,0"
        Style="{StaticResource BtnSelectOutline}" MinWidth="100"/>
<Button x:Name="BtnCrear" Content="Crear detail view"
        Style="{StaticResource BtnPrimary}" MinWidth="140"/>
"""


def _rgb(r, g, b, a=255):
    c = Color()
    c.A = a
    c.R = r
    c.G = g
    c.B = b
    return SolidColorBrush(c)


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


def _parse_positive_mm(text, default):
    try:
        v = float(_as_unicode(text).replace(u",", u".").strip())
        if v >= 50.0:
            return v
    except Exception:
        pass
    return float(default)


class _CrearDetailHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        win._execute_crear(uiapp)

    def GetName(self):
        return _TX_CREAR


class SeccionDetalleExtremoWindow(object):
    def __init__(self, uiapp, wall, parent_view, vft_detail, section_filter_text):
        self._uiapp = uiapp
        self._wall = wall
        self._parent_view = parent_view
        self._vft_detail = vft_detail
        self._section_filter_text = section_filter_text or u""
        self._preview = wall_section_preview_data(uiapp.ActiveUIDocument.Document, wall)

        xaml = build_simple_tool_xaml(
            title=_TOOL_TITLE,
            styles_xml=BIMTOOLS_DARK_STYLES_XML,
            body_xaml=_BODY_XAML,
            footer_actions_xaml=_FOOTER_ACTIONS,
            width=560,
            min_width=480,
            height=0,
            resize_mode=u"CanResize",
            size_to_content_height=True,
        )
        self._win = XamlReader.Parse(xaml)
        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_status = self._win.FindName(u"TxtStatus")
        self._txt_parent = self._win.FindName(u"TxtParent")
        self._txt_wall = self._win.FindName(u"TxtWall")
        self._txt_vft = self._win.FindName(u"TxtVft")
        self._canvas = self._win.FindName(u"SectionCanvas")
        self._rb_inicio = self._win.FindName(u"RbInicio")
        self._rb_termino = self._win.FindName(u"RbTermino")
        self._txt_far = self._win.FindName(u"TxtFarClip")
        self._txt_longitud = self._win.FindName(u"TxtLongitud")
        self._txt_nombre = self._win.FindName(u"TxtNombre")
        self._btn_crear = self._win.FindName(u"BtnCrear")

        if self._txt_longitud is not None:
            try:
                self._txt_longitud.Text = u"{0:.0f}".format(_LONGITUD_DETALLE_MM_DEF)
            except Exception:
                pass

        self._handler = _CrearDetailHandler(weakref.ref(self))
        self._ext_event = ExternalEvent.Create(self._handler)

        self._wire_events()
        self._fill_header()
        self._update_nombre()
        _prepare_window(self._win, uiapp)
        # Redibujar cuando el layout tenga tamaño
        self._win.ContentRendered += EventHandler(self._on_content_rendered)

    def _wire_events(self):
        from System.Windows import RoutedEventHandler

        self._win.FindName(u"BtnCrear").Click += RoutedEventHandler(self._on_crear)
        self._win.FindName(u"BtnCancel").Click += RoutedEventHandler(self._on_cancel)
        self._win.KeyDown += KeyEventHandler(self._on_key_down)
        self._win.Closed += EventHandler(self._on_closed)
        if self._rb_inicio is not None:
            self._rb_inicio.Checked += RoutedEventHandler(self._on_end_changed)
        if self._rb_termino is not None:
            self._rb_termino.Checked += RoutedEventHandler(self._on_end_changed)

    def _on_content_rendered(self, sender, args):
        try:
            self._draw_section_canvas()
        except Exception:
            pass

    def _on_key_down(self, sender, args):
        if args.Key == Key.Escape:
            self._win.Close()

    def _on_cancel(self, sender, args):
        self._win.Close()

    def _on_closed(self, sender, args):
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass

    def _on_end_changed(self, sender, args):
        self._update_nombre()
        try:
            self._draw_section_canvas()
        except Exception:
            pass

    def _on_crear(self, sender, args):
        self._set_status(u"Creando Detail View…")
        self._ext_event.Raise()

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _end_index(self):
        try:
            if self._rb_termino is not None and self._rb_termino.IsChecked:
                return 1
        except Exception:
            pass
        return 0

    def _fill_header(self):
        try:
            pv_name = _as_unicode(self._parent_view.Name)
        except Exception:
            pv_name = u"(vista)"
        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = (
                u"Corte horizontal mirando hacia abajo · extremo del muro en planta (confinamiento)."
            )
        if self._txt_parent is not None:
            self._txt_parent.Text = (
                u"Vista padre: {0}  ·  Section Filter: {1}"
            ).format(pv_name, self._section_filter_text or u"—")
        p = self._preview or {}
        if self._txt_wall is not None:
            self._txt_wall.Text = (
                u"Muro {0}  ·  {1}  ·  e={2:.0f} mm  ·  h={3:.0f} mm  ·  rebar={4}"
            ).format(
                p.get(u"marca", u"—"),
                p.get(u"tipo", u"—"),
                float(p.get(u"espesor_mm", 0) or 0),
                float(p.get(u"altura_mm", 0) or 0),
                int(p.get(u"n_rebar", 0) or 0),
            )
        vft_name = _view_family_type_display_name(self._vft_detail)
        if self._txt_vft is not None:
            self._txt_vft.Text = (
                u"Tipo Detail (automático): {0}"
            ).format(vft_name or u"—")

    def _update_nombre(self):
        if self._txt_nombre is None:
            return
        try:
            self._txt_nombre.Text = proposed_view_name(self._wall, self._end_index())
        except Exception:
            self._txt_nombre.Text = u"—"

    def _draw_section_canvas(self):
        canvas = self._canvas
        if canvas is None:
            return
        canvas.Children.Clear()
        try:
            cw = float(canvas.ActualWidth)
            ch = float(canvas.ActualHeight)
        except Exception:
            cw, ch = 0.0, 0.0
        if cw < 40 or ch < 40:
            try:
                parent = canvas.Parent
                cw = float(parent.ActualWidth) - 4
                ch = float(parent.ActualHeight) - 4
            except Exception:
                cw, ch = 500.0, 260.0
        if cw < 40:
            cw = 500.0
        if ch < 40:
            ch = 260.0

        p = self._preview or {}
        th_mm = max(float(p.get(u"espesor_mm", 200) or 200), 50.0)
        long_mm = max(float(p.get(u"longitud_detalle_mm", 1000) or 1000), 300.0)
        n_vert = int(p.get(u"n_vert_conf", 4) or 4)

        pad = 36.0
        avail_w = cw - 2 * pad
        avail_h = ch - 2 * pad - 28.0
        # Planta: Y = longitud hacia interior (arriba), X = espesor
        scale = min(avail_w / th_mm, avail_h / long_mm)
        # Ampliar espesor visualmente un poco si es muy estrecho
        wall_w = max(th_mm * scale, 56.0)
        wall_h = long_mm * scale
        if wall_h > avail_h:
            wall_h = avail_h
            scale = wall_h / long_mm
            wall_w = max(th_mm * scale, 56.0)
        x0 = (cw - wall_w) * 0.5
        y0 = pad + 10.0

        brush_wall = _rgb(42, 74, 92, 200)
        brush_border = _rgb(90, 122, 138)
        brush_rebar = _rgb(196, 92, 38)
        brush_dim = _rgb(138, 160, 181)
        brush_mark = _rgb(232, 184, 74)
        brush_zone = _rgb(232, 184, 74, 45)

        # Extremo abajo en pantalla; interior hacia arriba (BasisY)
        tip = u"INICIO" if self._end_index() == 0 else u"TÉRMINO"

        zone_h = min(wall_h * 0.35, wall_h)
        zone = Rectangle()
        zone.Width = wall_w
        zone.Height = zone_h
        zone.Fill = brush_zone
        zone.Stroke = brush_mark
        zone.StrokeThickness = 1.0
        Canvas.SetLeft(zone, x0)
        Canvas.SetTop(zone, y0 + wall_h - zone_h)
        canvas.Children.Add(zone)

        rect = Rectangle()
        rect.Width = wall_w
        rect.Height = wall_h
        rect.Fill = brush_wall
        rect.Stroke = brush_border
        rect.StrokeThickness = 1.5
        Canvas.SetLeft(rect, x0)
        Canvas.SetTop(rect, y0)
        canvas.Children.Add(rect)

        # Verticales de confinamiento en planta = círculos cerca del extremo (abajo)
        cover_x = max(8.0, wall_w * 0.18)
        cover_y = max(10.0, zone_h * 0.25)
        cols = 2
        rows = max(2, (n_vert + 1) // 2)
        for r in range(rows):
            for c in range(cols):
                xx = x0 + cover_x + c * max(wall_w - 2 * cover_x, 1.0)
                yy = y0 + wall_h - cover_y - r * (zone_h * 0.35)
                el = Ellipse()
                el.Width = 8
                el.Height = 8
                el.Fill = brush_rebar
                Canvas.SetLeft(el, xx - 4)
                Canvas.SetTop(el, yy - 4)
                canvas.Children.Add(el)

        # Flecha mirada hacia abajo (símbolo de sección)
        cx = x0 + wall_w + 32
        cy = y0 + wall_h * 0.5
        ring = Ellipse()
        ring.Width = 26
        ring.Height = 26
        ring.Stroke = brush_mark
        ring.StrokeThickness = 1.5
        ring.Fill = _rgb(7, 16, 24, 0)
        Canvas.SetLeft(ring, cx - 13)
        Canvas.SetTop(ring, cy - 13)
        canvas.Children.Add(ring)
        # Triángulo abajo = mirada
        arrow = Line()
        arrow.X1, arrow.Y1, arrow.X2, arrow.Y2 = cx, cy - 8, cx, cy + 10
        arrow.Stroke = brush_mark
        arrow.StrokeThickness = 2.0
        canvas.Children.Add(arrow)
        a1 = Line()
        a1.X1, a1.Y1, a1.X2, a1.Y2 = cx, cy + 10, cx - 5, cy + 4
        a1.Stroke = brush_mark
        a1.StrokeThickness = 2.0
        canvas.Children.Add(a1)
        a2 = Line()
        a2.X1, a2.Y1, a2.X2, a2.Y2 = cx, cy + 10, cx + 5, cy + 4
        a2.Stroke = brush_mark
        a2.StrokeThickness = 2.0
        canvas.Children.Add(a2)

        look = TextBlock()
        look.Text = u"mirada↓"
        look.Foreground = brush_mark
        look.FontSize = 10
        Canvas.SetLeft(look, cx - 16)
        Canvas.SetTop(look, cy + 16)
        canvas.Children.Add(look)

        lbl = TextBlock()
        lbl.Text = u"Planta · corte horizontal · confinamiento en {0}".format(tip)
        lbl.Foreground = brush_mark
        lbl.FontSize = 11
        Canvas.SetLeft(lbl, pad)
        Canvas.SetTop(lbl, ch - 20)
        canvas.Children.Add(lbl)

        lbl_th = TextBlock()
        lbl_th.Text = u"e={0:.0f}".format(th_mm)
        lbl_th.Foreground = brush_dim
        lbl_th.FontSize = 11
        Canvas.SetLeft(lbl_th, x0 + wall_w * 0.5 - 14)
        Canvas.SetTop(lbl_th, y0 - 16)
        canvas.Children.Add(lbl_th)

        end_lbl = TextBlock()
        end_lbl.Text = tip
        end_lbl.Foreground = brush_mark
        end_lbl.FontSize = 11
        Canvas.SetLeft(end_lbl, x0 + wall_w * 0.5 - 16)
        Canvas.SetTop(end_lbl, y0 + wall_h + 2)
        canvas.Children.Add(end_lbl)

        int_lbl = TextBlock()
        int_lbl.Text = u"↑ interior"
        int_lbl.Foreground = brush_dim
        int_lbl.FontSize = 10
        Canvas.SetLeft(int_lbl, x0 + wall_w + 4)
        Canvas.SetTop(int_lbl, y0)
        canvas.Children.Add(int_lbl)

    def _execute_crear(self, uiapp):
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            self._set_status(u"No hay documento activo.")
            return
        far = _parse_positive_mm(
            self._txt_far.Text if self._txt_far is not None else u"400",
            _FAR_CLIP_MM_DEF,
        )
        longitud = _parse_positive_mm(
            self._txt_longitud.Text if self._txt_longitud is not None else u"1000",
            _LONGITUD_DETALLE_MM_DEF,
        )
        ok, msg = ejecutar_crear_detail_extremo(
            uidoc,
            self._wall,
            self._end_index(),
            self._vft_detail,
            far_clip_mm=far,
            longitud_mm=longitud,
        )
        if ok:
            self._set_status(msg)
            try:
                self._win.Close()
            except Exception:
                pass
        else:
            self._set_status(u"Error: {0}".format(msg))
            try:
                show_message_dialog(
                    _DIALOG_TITLE,
                    instruction=u"No se pudo crear el Detail View.",
                    content=_as_unicode(msg),
                    ok_text=u"Entendido",
                    hwnd_revit=revit_main_hwnd(uiapp),
                    uiapp=uiapp,
                )
            except Exception:
                pass

    def Show(self):
        self._win.Show()


def _focus_existing():
    try:
        win = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        win = None
    if win is None:
        return False
    try:
        wrapper = win
        w = wrapper._win if hasattr(wrapper, u"_win") else wrapper
        if w is None:
            return False
        if w.WindowState == WindowState.Minimized:
            w.WindowState = WindowState.Normal
        w.Activate()
        w.Focus()
        try:
            show_message_dialog(
                _DIALOG_TITLE,
                instruction=u"La herramienta ya está en ejecución.",
                ok_text=u"Entendido",
            )
        except Exception:
            pass
        return True
    except Exception:
        try:
            AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
        except Exception:
            pass
        return False


def show_detalle_extremo_window(
    uiapp, wall, parent_view, vft_detail, section_filter_text
):
    if _focus_existing():
        return
    wrapper = SeccionDetalleExtremoWindow(
        uiapp,
        wall=wall,
        parent_view=parent_view,
        vft_detail=vft_detail,
        section_filter_text=section_filter_text,
    )
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, wrapper)
    except Exception:
        pass
    wrapper.Show()
