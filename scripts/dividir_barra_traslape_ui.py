# -*- coding: utf-8 -*-
"""UI WPF — dividir barra con traslape (esquema + selección de puntos en Revit)."""

from __future__ import print_function

import weakref

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPIUI")
clr.AddReference("RevitAPI")

from System import AppDomain, EventHandler, Action
from System.Windows import (
    RoutedEventHandler,
    FontWeights,
)
from System.Windows.Controls import (
    Button,
    Canvas,
    ComboBoxItem,
    ListBox,
    SelectionChangedEventHandler,
    StackPanel,
    TextBlock,
)
from System.Windows.Input import MouseButtonEventHandler
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Shapes import Line, Rectangle
from System.Windows.Threading import DispatcherPriority
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from dividir_barra_traslape_punto import (
    SPLICE_MODE_LABELS,
    SPLICE_MODES,
    SPLICE_SYMMETRIC,
    _DIALOG_TITLE,
    _DiagSession,
    _element_id_int,
    _exception_text,
    _validate_cuts_on_main,
    dividir_rebar_en_cortes,
    normalize_splice_mode,
    prepare_dividir_session,
    revit_pick_main_span_cut_mm,
    splice_overlap_zone_mm,
)

_SINGLETON_KEY = u"BIMTools.DividirBarraTraslape.ActiveWindow"
_CANVAS_W = 720.0
_CANVAS_H = 140.0
_MARGIN_X = 36.0
_BAR_H = 18.0

_COLOR_PREFIX = u"#64748b"
_COLOR_MAIN = u"#34d399"
_COLOR_SUFFIX = u"#64748b"
_COLOR_CUT = u"#f87171"
_COLOR_LAP = u"#fbbf24"


def _brush(hex_color, alpha=255):
    h = hex_color.lstrip(u"#")
    if len(h) == 6:
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
        return SolidColorBrush(Color.FromArgb(alpha, r, g, b))
    return SolidColorBrush(Color.FromRgb(200, 200, 200))


def _escape_xaml(text):
    try:
        s = unicode(text)
    except NameError:
        s = str(text)
    return (
        s.replace(u"&", u"&amp;")
        .replace(u"<", u"&lt;")
        .replace(u">", u"&gt;")
        .replace(u'"', u"&quot;")
    )


XAML = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="__TITLE__"
    Width="780" Height="580"
    WindowStartupLocation="CenterScreen"
    Background="Transparent"
    AllowsTransparency="True"
    FontFamily="Segoe UI"
    WindowStyle="None"
    ResizeMode="CanResizeWithGrip"
    MinWidth="640" MinHeight="480">
  <Window.Resources>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
  </Window.Resources>
  <Border CornerRadius="8" Background="#071018" BorderBrush="#21465C"
          BorderThickness="1" Padding="18,16">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <TextBlock Grid.Row="0" Text="__TITLE__" Foreground="#E8F4F8"
                 FontSize="15" FontWeight="Bold"/>
      <TextBlock x:Name="TxtSubtitle" Grid.Row="1" Margin="0,6,0,0"
                 Foreground="#95B8CC" FontSize="11" TextWrapping="Wrap"/>
      <TextBlock Grid.Row="2" Margin="0,10,0,6"
                 Text="Use «Seleccionar punto en Revit» sobre el vano (verde). También puede marcar en el esquema. Elija el tipo de solape antes de aplicar."
                 Foreground="#64748b" FontSize="10" TextWrapping="Wrap"/>
      <Border Grid.Row="3" Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="6" Padding="8" Margin="0,0,0,10">
        <Canvas x:Name="CnvBar" Width="__CW__" Height="__CH__"
                Background="#050E18" Cursor="Cross"/>
      </Border>
      <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,8">
        <TextBlock Text="Tipo de solape:" Foreground="#95B8CC"
                   FontSize="11" VerticalAlignment="Center" Margin="0,0,10,0"/>
        <ComboBox x:Name="CmbSplice" Style="{StaticResource Combo}"
                  Width="300" VerticalAlignment="Center"/>
      </StackPanel>
      <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="0,0,0,8">
        <Button x:Name="BtnPick" Content="Seleccionar punto en Revit"
                Style="{StaticResource BtnPrimary}" MinWidth="200"/>
        <Button x:Name="BtnClear" Content="Limpiar" Margin="10,0,0,0"
                Style="{StaticResource BtnSelectOutline}" MinWidth="80"/>
      </StackPanel>
      <StackPanel Grid.Row="6" Orientation="Horizontal" Margin="0,0,0,8">
        <TextBlock Text="Divisiones (mm en segmento mayor):" Foreground="#95B8CC"
                   FontSize="11" VerticalAlignment="Center" Margin="0,0,10,0"/>
        <ListBox x:Name="LstCuts" MinWidth="280" MaxHeight="72"
                 Background="#050E18" Foreground="#E8F4F8" BorderBrush="#21465C"/>
      </StackPanel>
      <Grid Grid.Row="7">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="TxtStatus" Grid.Column="0"
                   Foreground="#64748b" FontSize="10" TextWrapping="Wrap"
                   VerticalAlignment="Center"/>
        <Button x:Name="BtnApply" Grid.Column="1" Content="Aplicar divisiones"
                Style="{StaticResource BtnPrimary}" MinWidth="160" Margin="0,0,8,0"/>
        <Button x:Name="BtnClose" Grid.Column="2" Content="Cerrar"
                Style="{StaticResource BtnSelectOutline}" MinWidth="90"/>
      </Grid>
    </Grid>
  </Border>
</Window>
"""


def _singleton_get():
    try:
        return AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        return None


def _singleton_set(win):
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass


def _singleton_clear():
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
    except Exception:
        pass


def _try_focus_existing():
    win = _singleton_get()
    if win is None:
        return False
    try:
        if not win.IsLoaded:
            _singleton_clear()
            return False
        if win.WindowState == 2:
            win.WindowState = 0
        win.Activate()
        win.Topmost = True
        win.Topmost = False
        TaskDialog.Show(_DIALOG_TITLE, u"La herramienta ya esta en ejecucion.")
        return True
    except Exception:
        _singleton_clear()
        return False


class _AplicarDivisionHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def GetName(self):
        return u"DividirBarraTraslape.Aplicar"

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        try:
            win.apply_in_revit(uiapp)
        except Exception as ex:
            win._ui_set_status(_exception_text(ex))


class _PickPuntoHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def GetName(self):
        return u"DividirBarraTraslape.PickPunto"

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        doc = None
        try:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is None:
                win._ui_set_status(u"No hay documento activo.")
                return
            doc = uidoc.Document
            rebar = doc.GetElement(win._rebar_id)
            if rebar is None:
                win._ui_set_status(u"La barra ya no existe en el modelo.")
                return
            try:
                win._win.Hide()
            except Exception:
                pass
            cut_mm, err = revit_pick_main_span_cut_mm(uiapp, rebar)
            if cut_mm is not None:
                win._ui_add_cut(cut_mm)
            elif err:
                win._ui_set_status(err)
            else:
                win._ui_set_status(u"Selección cancelada.")
        except OperationCanceledException:
            win._ui_set_status(u"Selección cancelada.")
        except Exception as ex:
            win._ui_set_status(_exception_text(ex))
        finally:
            try:
                win._win.Show()
                win._win.Activate()
            except Exception:
                pass


class DividirBarraWindow(object):
    def __init__(self, uiapp, session):
        self._uiapp = uiapp
        self._session = session or {}
        self._preview = session.get(u"preview") or {}
        self._rebar_id = session.get(u"rebar_id")
        self._lap_mm = float(session.get(u"lap_mm") or 0.0)
        self._cuts_mm = list(session.get(u"suggested_cuts_mm") or [])
        self._splice_mode = normalize_splice_mode(
            session.get(u"splice_mode") or SPLICE_SYMMETRIC
        )
        self._scale = 1.0
        self._main_x0 = 0.0
        self._main_x1 = 0.0
        self._bar_y = _CANVAS_H * 0.5
        self._margin_x = _MARGIN_X

        xaml = (
            XAML.replace(u"__TITLE__", _escape_xaml(_DIALOG_TITLE))
            .replace(u"__CW__", unicode(int(_CANVAS_W)))
            .replace(u"__CH__", unicode(int(_CANVAS_H)))
        )
        self._win = XamlReader.Parse(xaml)
        self._cnv = self._win.FindName(u"CnvBar")
        self._lst = self._win.FindName(u"LstCuts")
        self._cmb = self._win.FindName(u"CmbSplice")
        self._txt_sub = self._win.FindName(u"TxtSubtitle")
        self._txt_status = self._win.FindName(u"TxtStatus")

        self._handler_apply = _AplicarDivisionHandler(weakref.ref(self))
        self._handler_pick = _PickPuntoHandler(weakref.ref(self))
        self._ext_apply = ExternalEvent.Create(self._handler_apply)
        self._ext_pick = ExternalEvent.Create(self._handler_pick)

        self._populate_splice_combo()
        self._wire_events()
        self._update_subtitle()
        self._refresh_ui()

    def _populate_splice_combo(self):
        if self._cmb is None:
            return
        self._cmb.Items.Clear()
        selected_idx = 0
        for i, mode in enumerate(SPLICE_MODES):
            item = ComboBoxItem()
            item.Content = SPLICE_MODE_LABELS.get(mode, mode)
            item.Tag = mode
            self._cmb.Items.Add(item)
            if mode == self._splice_mode:
                selected_idx = i
        self._cmb.SelectedIndex = selected_idx

    def _current_splice_mode(self):
        if self._cmb is None:
            return self._splice_mode
        try:
            item = self._cmb.SelectedItem
            if item is not None and getattr(item, u"Tag", None):
                return normalize_splice_mode(item.Tag)
        except Exception:
            pass
        return normalize_splice_mode(self._splice_mode)

    def _wire_events(self):
        self._win.FindName(u"BtnClose").Click += RoutedEventHandler(self._on_close)
        self._win.FindName(u"BtnApply").Click += RoutedEventHandler(self._on_apply)
        self._win.FindName(u"BtnClear").Click += RoutedEventHandler(self._on_clear)
        self._win.FindName(u"BtnPick").Click += RoutedEventHandler(self._on_pick)
        self._cnv.MouseLeftButtonDown += MouseButtonEventHandler(self._on_canvas_left)
        self._cnv.MouseRightButtonDown += MouseButtonEventHandler(self._on_canvas_right)
        self._win.Closed += EventHandler(self._on_closed)
        if self._cmb is not None:
            self._cmb.SelectionChanged += SelectionChangedEventHandler(
                self._on_splice_changed
            )

    def _invoke_ui(self, action):
        try:
            self._win.Dispatcher.Invoke(
                DispatcherPriority.Normal, Action(action)
            )
        except Exception:
            try:
                action()
            except Exception:
                pass

    def _ui_add_cut(self, main_mm):
        def _do():
            self._try_add_cut(main_mm)
        self._invoke_ui(_do)

    def _ui_set_status(self, msg):
        def _do():
            self.set_status(msg)
        self._invoke_ui(_do)

    def _update_subtitle(self):
        p = self._preview
        n_pos = self._session.get(u"n_pos") or 1
        rule = self._session.get(u"layout_rule") or u"?"
        parts = [
            u"Rebar Id {0}".format(_element_id_int(self._rebar_id)),
            u"Ø {0} mm".format(self._session.get(u"diameter_mm") or u"?"),
            u"Traslape {0:.0f} mm".format(self._lap_mm),
            u"Segmento mayor {0:.0f} mm".format(p.get(u"main_length_mm") or 0),
        ]
        if n_pos > 1:
            parts.append(u"Layout {0} ({1} pos.)".format(rule, n_pos))
        if self._txt_sub is not None:
            self._txt_sub.Text = u" · ".join(parts)

    def set_status(self, msg):
        if self._txt_status is not None:
            self._txt_status.Text = msg or u""

    def _layout_metrics(self):
        preview = self._preview
        segments = preview.get(u"segments") or []
        total_mm = float(preview.get(u"total_length_mm") or 0.0)
        if total_mm <= 1e-6:
            return [], None, 1.0
        usable = _CANVAS_W - 2.0 * _MARGIN_X
        scale = usable / total_mm
        x = _MARGIN_X
        bounds = []
        main_bounds = None
        for seg in segments:
            L = float(seg.get(u"length_mm") or 0.0)
            w = L * scale
            role = seg.get(u"role") or u"main"
            x0, x1 = x, x + w
            bounds.append((x0, x1, role, L))
            if role == u"main":
                main_bounds = (x0, x1)
            x = x1
        return bounds, main_bounds, scale

    def _canvas_x_to_main_mm(self, canvas_x):
        preview = self._preview
        prefix_mm = float(preview.get(u"prefix_length_mm") or 0.0)
        dist_bar = (float(canvas_x) - self._margin_x) / self._scale
        return dist_bar - prefix_mm

    def _main_mm_to_canvas_x(self, main_mm):
        preview = self._preview
        prefix_mm = float(preview.get(u"prefix_length_mm") or 0.0)
        return self._margin_x + (prefix_mm + float(main_mm)) * self._scale

    def _redraw_canvas(self):
        if self._cnv is None:
            return
        self._cnv.Children.Clear()
        bounds, main_bounds, scale = self._layout_metrics()
        self._scale = scale
        if main_bounds:
            self._main_x0, self._main_x1 = main_bounds
        y0 = self._bar_y - _BAR_H * 0.5
        for x0, x1, role, _L in bounds:
            if role == u"main":
                fill = _brush(_COLOR_MAIN)
            elif role == u"prefix":
                fill = _brush(_COLOR_PREFIX)
            else:
                fill = _brush(_COLOR_SUFFIX)
            rect = Rectangle()
            rect.Width = max(1.0, x1 - x0)
            rect.Height = _BAR_H
            rect.Fill = fill
            rect.Stroke = _brush(u"#1e293b")
            rect.StrokeThickness = 1.0
            Canvas.SetLeft(rect, x0)
            Canvas.SetTop(rect, y0)
            self._cnv.Children.Add(rect)

        mode = self._current_splice_mode()
        main_len = float(self._preview.get(u"main_length_mm") or 0.0)
        for c in sorted(self._cuts_mm):
            cx = self._main_mm_to_canvas_x(c)
            a_mm, b_mm = splice_overlap_zone_mm(c, self._lap_mm, mode)
            a_mm = max(0.0, a_mm)
            b_mm = min(main_len, b_mm) if main_len > 0 else b_mm
            x_a = self._main_mm_to_canvas_x(a_mm)
            x_b = self._main_mm_to_canvas_x(b_mm)
            lap_w = max(2.0, x_b - x_a)
            lap_rect = Rectangle()
            lap_rect.Width = lap_w
            lap_rect.Height = _BAR_H + 8.0
            lap_rect.Fill = _brush(_COLOR_LAP, 48)
            lap_rect.Stroke = _brush(_COLOR_LAP, 100)
            lap_rect.StrokeThickness = 1.0
            Canvas.SetLeft(lap_rect, x_a)
            Canvas.SetTop(lap_rect, y0 - 4.0)
            self._cnv.Children.Add(lap_rect)

            cut_ln = Line()
            cut_ln.X1 = cx
            cut_ln.X2 = cx
            cut_ln.Y1 = y0 - 10.0
            cut_ln.Y2 = y0 + _BAR_H + 10.0
            cut_ln.Stroke = _brush(_COLOR_CUT)
            cut_ln.StrokeThickness = 2.0
            self._cnv.Children.Add(cut_ln)

            lbl = TextBlock()
            lbl.Text = u"{:.0f}".format(c)
            lbl.Foreground = _brush(_COLOR_CUT)
            lbl.FontSize = 10.0
            lbl.FontWeight = FontWeights.Bold
            Canvas.SetLeft(lbl, cx - 14.0)
            Canvas.SetTop(lbl, y0 + _BAR_H + 12.0)
            self._cnv.Children.Add(lbl)

    def _refresh_list(self):
        if self._lst is None:
            return
        self._lst.Items.Clear()
        for c in sorted(self._cuts_mm):
            self._lst.Items.Add(u"{:.0f} mm".format(c))

    def _refresh_ui(self):
        self._redraw_canvas()
        self._refresh_list()
        n = len(self._cuts_mm)
        mode = self._current_splice_mode()
        mode_lbl = SPLICE_MODE_LABELS.get(mode, mode)
        self.set_status(
            u"{0} división(es) → {1} tramo(s) · solape: {2}.".format(
                n, n + 1 if n else 0, mode_lbl
            )
        )

    def _on_splice_changed(self, sender, args):
        self._splice_mode = self._current_splice_mode()
        if self._cuts_mm:
            main_len = float(self._preview.get(u"main_length_mm") or 0.0)
            ok, err = _validate_cuts_on_main(
                self._cuts_mm, main_len, self._lap_mm, splice_mode=self._splice_mode
            )
            if not ok:
                self._cuts_mm = []
                self._refresh_ui()
                self.set_status(
                    u"Cortes no válidos para este solape y se limpiaron. {0}".format(
                        err or u""
                    )
                )
                return
        self._refresh_ui()

    def _try_add_cut(self, main_mm):
        main_len = float(self._preview.get(u"main_length_mm") or 0.0)
        trial = sorted(self._cuts_mm + [float(main_mm)])
        ok, err = _validate_cuts_on_main(
            trial, main_len, self._lap_mm, splice_mode=self._current_splice_mode()
        )
        if not ok:
            self.set_status(err)
            return
        self._cuts_mm = trial
        self._refresh_ui()

    def _on_canvas_left(self, sender, args):
        pos = args.GetPosition(self._cnv)
        main_mm = self._canvas_x_to_main_mm(pos.X)
        main_len = float(self._preview.get(u"main_length_mm") or 0.0)
        if main_mm < 0 or main_mm > main_len:
            self.set_status(u"Clic solo sobre el segmento mayor (verde).")
            return
        self._try_add_cut(main_mm)

    def _on_canvas_right(self, sender, args):
        if not self._cuts_mm:
            return
        pos = args.GetPosition(self._cnv)
        main_mm = self._canvas_x_to_main_mm(pos.X)
        nearest = min(self._cuts_mm, key=lambda c: abs(c - main_mm))
        self._cuts_mm = [c for c in self._cuts_mm if abs(c - nearest) > 0.5]
        self._refresh_ui()

    def _on_pick(self, sender, args):
        self.set_status(u"Seleccione un punto en Revit sobre el vano…")
        self._ext_pick.Raise()

    def _on_clear(self, sender, args):
        self._cuts_mm = []
        self._refresh_ui()

    def _on_apply(self, sender, args):
        if not self._cuts_mm:
            self.set_status(u"Añada al menos una división.")
            return
        self.set_status(u"Aplicando en Revit…")
        self._ext_apply.Raise()

    def apply_in_revit(self, uiapp):
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            self._ui_set_status(u"No hay documento activo.")
            return
        doc = uidoc.Document
        rebar = doc.GetElement(self._rebar_id)
        if rebar is None:
            self._ui_set_status(u"La barra ya no existe en el modelo.")
            return
        mode = self._current_splice_mode()
        diag = _DiagSession()
        ok, msg, ids = dividir_rebar_en_cortes(
            doc,
            rebar,
            self._cuts_mm,
            lap_mm=self._lap_mm,
            diag=diag,
            splice_mode=mode,
        )
        if ok:
            ids_txt = u", ".join(
                unicode(_element_id_int(i)) for i in (ids or [])
            )
            TaskDialog.Show(
                _DIALOG_TITLE,
                u"{0}\n\nNuevas barras (Id): {1}".format(msg, ids_txt),
            )
            try:
                self._win.Close()
            except Exception:
                pass
        else:
            self._ui_set_status(msg)

    def _on_close(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass

    def _on_closed(self, sender, args):
        _singleton_clear()

    def show(self):
        _singleton_set(self._win)
        self._win.Show()


def show_dividir_barra_window(uiapp, rebar):
    if _try_focus_existing():
        return
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(_DIALOG_TITLE, u"No hay documento activo.")
        return
    ok, err, session = prepare_dividir_session(uidoc.Document, rebar)
    if not ok:
        TaskDialog.Show(_DIALOG_TITLE, err or u"No se pudo preparar la barra.")
        return
    try:
        win = DividirBarraWindow(uiapp, session)
        win.show()
    except Exception as ex:
        _singleton_clear()
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"Error al abrir la ventana:\n\n{0}".format(_exception_text(ex)),
        )
