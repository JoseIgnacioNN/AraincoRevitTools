# -*- coding: utf-8 -*-
"""UI WPF — Eliminar segmento rebar (alzado clicable, multi-barra)."""

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

from System import Action, AppDomain, Byte, Double
from System.Collections.Generic import List as DotNetList
from System.Windows import (
    FontWeights,
    Point,
    RoutedEventHandler,
    SizeToContent,
    WindowStartupLocation,
    WindowState,
)
from System.Windows.Controls import Canvas, TextBlock
from System.Windows.Input import (
    MouseButton,
    MouseButtonEventHandler,
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
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler

from bimtools_ui_tokens import ACCENT_PRIMARY, BTN_MANUAL, FG_BODY, FG_MUTED, FONT_SIZE_HINT
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_wpf_shell import build_simple_tool_xaml
from dividir_rebar_punto_core import resolve_active_model_view
from dividir_rebar_punto_geom import compute_canvas_mapping, map_uv_to_canvas_px
from eliminar_segmento_rebar import _ALREADY_RUNNING, _TITULO, _as_unicode, mostrar_aviso
from eliminar_segmento_rebar_core import (
    apply_remove_segments,
    build_elevation_session,
    collect_all_uv,
    is_end_segment,
    match_segments,
    pick_rebars,
    segment_letter,
)
from revit_wpf_window_position import (
    bind_maximize_wpf_on_revit_monitor,
    preposition_wpf_window_on_work_area,
    revit_main_hwnd,
    _monitor_work_area_px,
    _primary_work_area_px,
)

_SINGLETON_KEY = u"Arainco_EliminarSegmentoRebar_UI"
_HOLD_KEY = u"Arainco_EliminarSegmentoRebar_HOLD"

_WINDOW_W = 1100
_WINDOW_H = 780
_WINDOW_MIN_W = 860
_WINDOW_MIN_H = 560
_CANVAS_MIN_W = 640.0
_CANVAS_MIN_H = 420.0
_MARGIN = 36.0
_HIT_PX = 12.0
_ZOOM_MIN = 0.5
_ZOOM_MAX = 8.0
_ZOOM_STEP = 1.15

_COLOR_BAR = u"#3d7a94"
_COLOR_HOVER = ACCENT_PRIMARY
_COLOR_SEL = u"#f0a060"
_COLOR_GONE = u"#5a3a3a"
_COLOR_KEEP = u"#4ade80"
_COLOR_CONCRETE_EDGE = u"#9ca3a8"
_COLOR_CONCRETE_FILL = (138, 142, 148, 120)

_BRUSH_CACHE = {}


def _brush(hex_or_rgba):
    key = hex_or_rgba
    if key in _BRUSH_CACHE:
        return _BRUSH_CACHE[key]
    if isinstance(hex_or_rgba, tuple):
        r, g, b = hex_or_rgba[0], hex_or_rgba[1], hex_or_rgba[2]
        a = hex_or_rgba[3] if len(hex_or_rgba) > 3 else 255
        br = SolidColorBrush(Color.FromArgb(Byte(a), Byte(r), Byte(g), Byte(b)))
    else:
        h = _as_unicode(hex_or_rgba).lstrip(u"#")
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
        br = SolidColorBrush(Color.FromRgb(Byte(r), Byte(g), Byte(b)))
    try:
        br.Freeze()
    except Exception:
        pass
    _BRUSH_CACHE[key] = br
    return br


def _window_alive(win):
    try:
        return win is not None and bool(win.IsLoaded)
    except Exception:
        return False


def _existing_ctrl():
    try:
        ctrl = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        ctrl = None
    if ctrl is not None and _window_alive(getattr(ctrl, u"_win", None)):
        return ctrl
    try:
        from System.Windows import Application

        app = Application.Current
        if app is None:
            return None
        for ww in app.Windows:
            try:
                txt = ww.FindName(u"TxtTitle")
                if txt is not None and _as_unicode(txt.Text) == _TITULO:
                    if _window_alive(ww):
                        class _Wrap(object):
                            def __init__(self, w):
                                self._win = w

                        return _Wrap(ww)
            except Exception:
                continue
    except Exception:
        pass
    return None


def _set_singleton(ctrl):
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, ctrl)
    except Exception:
        pass


def _clear_singleton():
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
    except Exception:
        pass


def _set_hold(ctrl):
    try:
        AppDomain.CurrentDomain.SetData(_HOLD_KEY, ctrl)
    except Exception:
        pass


def _clear_hold():
    try:
        AppDomain.CurrentDomain.SetData(_HOLD_KEY, None)
    except Exception:
        pass


def _activate_existing(ctrl, uiapp):
    win = getattr(ctrl, u"_win", None)
    if not _window_alive(win):
        _clear_singleton()
        return False
    try:
        if win.WindowState == WindowState.Minimized:
            win.WindowState = WindowState.Normal
    except Exception:
        pass
    try:
        win.Activate()
        win.Focus()
    except Exception:
        pass
    mostrar_aviso(uiapp, _ALREADY_RUNNING)
    return True


def _resolve_manual_path():
    candidates = []
    try:
        import bimtools_paths

        pb = bimtools_paths.get_pushbutton_dir()
        if pb:
            candidates.append(os.path.join(pb, u"manual_usuario.html"))
    except Exception:
        pass
    try:
        ext_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        for root, dirs, files in os.walk(ext_dir):
            base = os.path.basename(root)
            if base.endswith(u"EliminarSegmentoRebar.pushbutton"):
                candidates.append(os.path.join(root, u"manual_usuario.html"))
                dirs[:] = []
                continue
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
    for path in candidates:
        try:
            ap = os.path.normpath(os.path.abspath(path))
        except Exception:
            continue
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
        mostrar_aviso(uiapp, u"No se pudo abrir el manual.", _as_unicode(ex))


def _clear_selection(uidoc):
    if uidoc is None:
        return
    try:
        uidoc.Selection.SetElementIds(DotNetList[ElementId]())
    except Exception:
        pass


def _dist_point_seg(px, py, x0, y0, x1, y1):
    dx = x1 - x0
    dy = y1 - y0
    L2 = dx * dx + dy * dy
    if L2 < 1e-9:
        return ((px - x0) ** 2 + (py - y0) ** 2) ** 0.5
    t = ((px - x0) * dx + (py - y0) * dy) / L2
    if t < 0.0:
        t = 0.0
    elif t > 1.0:
        t = 1.0
    qx = x0 + t * dx
    qy = y0 + t * dy
    return ((px - qx) ** 2 + (py - qy) ** 2) ** 0.5


def _build_xaml():
    body = u"""
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <TextBlock x:Name="TxtMeta" Foreground="#95B8CC" FontSize="11"
             TextWrapping="Wrap" Margin="0,0,0,8"/>
  <Border x:Name="BdrCanvasHost" Grid.Row="1"
          Background="#050e18" BorderBrush="#1a3a4d" BorderThickness="1"
          CornerRadius="4" ClipToBounds="True" MinHeight="360">
    <Canvas x:Name="CnvElev" Background="#050e18"
            HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>
  </Border>
  <TextBlock x:Name="TxtLegend" Grid.Row="2" Margin="0,8,0,0"
             Foreground="#64748b" FontSize="10" TextWrapping="Wrap"
             Text="Clic en un tramo del alzado para seleccionarlo. Solo extremos (A o último) mantienen un sketch continuo. Rueda: zoom · botón medio: desplazar."/>
</Grid>
"""
    footer_leading = (
        u'<Button x:Name="BtnManual" Content="Manual" '
        u'Style="{{StaticResource BtnSelectOutline}}" '
        u'Background="{bg}" ToolTip="Abrir manual de usuario"/>'
    ).format(bg=BTN_MANUAL)
    footer_actions = u"""
<Button x:Name="BtnSelect" Content="Seleccionar barras"
        Style="{StaticResource BtnSelectOutline}" Margin="0,0,8,0"/>
<Button x:Name="BtnDelete" Content="Eliminar segmento"
        Style="{StaticResource BtnPrimary}" Margin="0,0,8,0" IsEnabled="False"/>
<Button x:Name="BtnClose" Content="Cerrar"
        Style="{StaticResource BtnSelectOutline}"/>
"""
    return build_simple_tool_xaml(
        title=_TITULO,
        styles_xml=BIMTOOLS_DARK_STYLES_XML,
        body_xaml=body,
        footer_leading_xaml=footer_leading,
        footer_actions_xaml=footer_actions,
        width=_WINDOW_W,
        min_width=_WINDOW_MIN_W,
        height=_WINDOW_H,
        min_height=_WINDOW_MIN_H,
        resize_mode=u"CanResize",
    )


class _PickHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def GetName(self):
        return u"EliminarSegmentoRebar.Pick"

    def Execute(self, uiapp):
        win = None
        try:
            win = self._window_ref() if self._window_ref is not None else None
        except Exception:
            win = None
        if win is None:
            try:
                win = AppDomain.CurrentDomain.GetData(_HOLD_KEY)
            except Exception:
                win = None
        if win is not None:
            win.pick_in_revit(uiapp)


class _ApplyHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def GetName(self):
        return u"EliminarSegmentoRebar.Apply"

    def Execute(self, uiapp):
        win = None
        try:
            win = self._window_ref() if self._window_ref is not None else None
        except Exception:
            win = None
        if win is None:
            try:
                win = AppDomain.CurrentDomain.GetData(_HOLD_KEY)
            except Exception:
                win = None
        try:
            if win is not None:
                win.apply_in_revit(uiapp)
        finally:
            _clear_hold()


class EliminarSegmentoWindow(object):
    def __init__(self, uiapp):
        self._uiapp = uiapp
        self._win = XamlReader.Parse(_build_xaml())
        self._session = None
        self._selected = []
        self._hover = None
        self._mapping = None
        self._zoom = 1.0
        self._pan_x = 0.0
        self._pan_y = 0.0
        self._panning = False
        self._pan_last = None
        self._canvas_w = float(_CANVAS_MIN_W)
        self._canvas_h = float(_CANVAS_MIN_H)
        self._busy = False

        self._txt_subtitle = self._win.FindName(u"TxtSubtitle")
        self._txt_meta = self._win.FindName(u"TxtMeta")
        self._txt_status = self._win.FindName(u"TxtStatus")
        self._txt_legend = self._win.FindName(u"TxtLegend")
        self._cnv = self._win.FindName(u"CnvElev")
        self._host = self._win.FindName(u"BdrCanvasHost")
        self._btn_manual = self._win.FindName(u"BtnManual")
        self._btn_select = self._win.FindName(u"BtnSelect")
        self._btn_delete = self._win.FindName(u"BtnDelete")
        self._btn_close = self._win.FindName(u"BtnClose")

        if self._txt_subtitle is not None:
            self._txt_subtitle.Text = (
                u"Clic en un tramo del alzado para marcarlo y eliminarlo "
                u"en las barras equivalentes."
            )

        href = weakref.ref(self)
        self._handler_pick = _PickHandler(href)
        self._handler_apply = _ApplyHandler(href)
        self._ext_pick = ExternalEvent.Create(self._handler_pick)
        self._ext_apply = ExternalEvent.Create(self._handler_apply)

        if self._cnv is not None:
            try:
                RenderOptions.SetEdgeMode(self._cnv, EdgeMode.Unspecified)
            except Exception:
                pass
        self._wire()
        self._set_status(u"Listo · sin selección")
        self._update_meta()
        self._redraw()

    def _wire(self):
        if self._btn_manual is not None:
            self._btn_manual.Click += RoutedEventHandler(self._on_manual)
        if self._btn_select is not None:
            self._btn_select.Click += RoutedEventHandler(self._on_select)
        if self._btn_delete is not None:
            self._btn_delete.Click += RoutedEventHandler(self._on_delete)
        if self._btn_close is not None:
            self._btn_close.Click += RoutedEventHandler(self._on_close)
        self._win.Closed += self._on_closed
        if self._cnv is not None:
            self._cnv.MouseLeftButtonDown += MouseButtonEventHandler(self._on_canvas_down)
            self._cnv.MouseMove += MouseEventHandler(self._on_canvas_move)
            self._cnv.MouseLeave += MouseEventHandler(self._on_canvas_leave)
            self._cnv.MouseWheel += MouseWheelEventHandler(self._on_wheel)
            self._cnv.MouseDown += MouseButtonEventHandler(self._on_mouse_down)
            self._cnv.MouseUp += MouseButtonEventHandler(self._on_mouse_up)
        if self._host is not None:
            self._host.SizeChanged += self._on_host_size

    def _set_status(self, text):
        if self._txt_status is not None:
            self._txt_status.Text = _as_unicode(text)

    def _update_meta(self):
        if self._txt_meta is None:
            return
        if not self._session:
            self._txt_meta.Text = u"Sin barras. Pulse «Seleccionar barras» o seleccione Rebar en la vista y vuelva a abrir."
            return
        bars = list(self._session.get(u"bars") or [])
        skipped = list(self._session.get(u"skipped") or [])
        n_seg = 0
        for b in bars:
            n_seg += int(b.get(u"n_segments") or 0)
        msg = u"{0} barra(s) · {1} tramo(s) en alzado.".format(len(bars), n_seg)
        if skipped:
            msg = msg + u" Omitidas: {0}.".format(len(skipped))
        if self._selected:
            letters = []
            for bi, si in self._selected:
                letters.append(segment_letter(si))
            uniq = []
            for L in letters:
                if L not in uniq:
                    uniq.append(L)
            msg = msg + u" Selección: tramo {0} en {1} barra(s).".format(
                u"/".join(uniq), len(self._selected)
            )
        self._txt_meta.Text = msg

    def _selection_valid(self):
        if not self._session or not self._selected:
            return False
        bars = list(self._session.get(u"bars") or [])
        for bi, si in self._selected:
            if bi < 0 or bi >= len(bars):
                return False
            n = int(bars[bi].get(u"n_segments") or 0)
            if not is_end_segment(si, n):
                return False
        return True

    def _update_delete_enabled(self):
        if self._btn_delete is not None:
            self._btn_delete.IsEnabled = bool(self._selection_valid())

    def _map_uv(self, u, v):
        if not self._mapping:
            return 0.0, 0.0
        x, y = map_uv_to_canvas_px(u, v, self._mapping)
        return x * self._zoom + self._pan_x, y * self._zoom + self._pan_y

    def _fit_mapping(self):
        pts = collect_all_uv(self._session)
        self._mapping = compute_canvas_mapping(
            pts,
            self._canvas_w,
            self._canvas_h,
            _MARGIN,
            swap_uv=False,
            flip_v=True,
        )

    def _sync_canvas_size(self):
        w = float(_CANVAS_MIN_W)
        h = float(_CANVAS_MIN_H)
        try:
            if self._host is not None:
                aw = float(self._host.ActualWidth or 0.0)
                ah = float(self._host.ActualHeight or 0.0)
                if aw > 40.0:
                    w = aw
                if ah > 40.0:
                    h = ah
        except Exception:
            pass
        self._canvas_w = w
        self._canvas_h = h
        if self._cnv is not None:
            try:
                self._cnv.Width = w
                self._cnv.Height = h
            except Exception:
                pass

    def _smooth_pen(self, shape):
        try:
            shape.StrokeStartLineCap = PenLineCap.Round
            shape.StrokeEndLineCap = PenLineCap.Round
            shape.StrokeLineJoin = PenLineJoin.Round
        except Exception:
            pass

    def _add_polyline(self, pts, stroke, thickness, dash=None, opacity=1.0):
        if self._cnv is None or len(pts) < 2:
            return
        pl = Polyline()
        pc = PointCollection()
        for x, y in pts:
            pc.Add(Point(float(x), float(y)))
        pl.Points = pc
        pl.Stroke = _brush(stroke)
        pl.StrokeThickness = float(thickness)
        pl.IsHitTestVisible = False
        if dash:
            try:
                from System.Windows.Media import DoubleCollection

                dc = DoubleCollection()
                for d in dash:
                    dc.Add(float(d))
                pl.StrokeDashArray = dc
            except Exception:
                pass
        try:
            pl.Opacity = float(opacity)
        except Exception:
            pass
        self._smooth_pen(pl)
        self._cnv.Children.Add(pl)

    def _add_label(self, x, y, text, color):
        tb = TextBlock()
        tb.Text = _as_unicode(text)
        tb.Foreground = _brush(color)
        tb.FontFamily = FontFamily(u"Segoe UI")
        tb.FontSize = 11
        tb.FontWeight = FontWeights.Bold
        tb.IsHitTestVisible = False
        Canvas.SetLeft(tb, float(x) - 5.0)
        Canvas.SetTop(tb, float(y) - 16.0)
        self._cnv.Children.Add(tb)

    def _selected_set(self):
        return set((int(a), int(b)) for a, b in (self._selected or []))

    def _redraw(self):
        if self._cnv is None:
            return
        self._cnv.Children.Clear()
        bg = Rectangle()
        bg.Width = self._canvas_w
        bg.Height = self._canvas_h
        bg.Fill = _brush(u"#050e18")
        bg.IsHitTestVisible = False
        self._cnv.Children.Add(bg)

        if not self._session:
            empty = TextBlock()
            empty.Text = u"Seleccione barras para dibujar el alzado."
            empty.Foreground = _brush(u"#64748b")
            empty.FontSize = 13
            Canvas.SetLeft(empty, 24.0)
            Canvas.SetTop(empty, self._canvas_h * 0.45)
            self._cnv.Children.Add(empty)
            return

        self._fit_mapping()
        sel = self._selected_set()
        hover = self._hover

        for rect in self._session.get(u"context_fill_rects_uv") or []:
            try:
                u0, u1, v0, v1 = [float(x) for x in rect]
            except Exception:
                continue
            corners = [
                self._map_uv(u0, v0),
                self._map_uv(u1, v0),
                self._map_uv(u1, v1),
                self._map_uv(u0, v1),
                self._map_uv(u0, v0),
            ]
            self._add_polyline(corners, _COLOR_CONCRETE_EDGE, 0.6, opacity=0.55)

        for pl in self._session.get(u"context_polylines_uv") or []:
            pts = [self._map_uv(p[0], p[1]) for p in (pl or [])]
            self._add_polyline(pts, _COLOR_CONCRETE_EDGE, 0.7, opacity=0.7)

        bars = list(self._session.get(u"bars") or [])
        for bi, bar in enumerate(bars):
            n = int(bar.get(u"n_segments") or 0)
            for seg in bar.get(u"segments") or []:
                si = int(seg.get(u"idx") or 0)
                uv = list(seg.get(u"uv") or [])
                if len(uv) < 2:
                    continue
                pts = [self._map_uv(p[0], p[1]) for p in uv]
                key = (bi, si)
                stroke = _COLOR_BAR
                thick = 2.6
                dash = None
                opacity = 1.0
                if key in sel:
                    if is_end_segment(si, n):
                        stroke = _COLOR_SEL
                        thick = 4.4
                    else:
                        stroke = _COLOR_GONE
                        thick = 3.0
                        dash = (6, 4)
                        opacity = 0.7
                elif hover == key:
                    stroke = _COLOR_HOVER
                    thick = 3.4
                self._add_polyline(pts, stroke, thick, dash=dash, opacity=opacity)
                mid = pts[len(pts) // 2]
                self._add_label(mid[0], mid[1], segment_letter(si), u"#e8f4f8")

    def _hit_test(self, px, py):
        if not self._session:
            return None
        best = None
        best_d = _HIT_PX + 1.0
        bars = list(self._session.get(u"bars") or [])
        for bi, bar in enumerate(bars):
            for seg in bar.get(u"segments") or []:
                uv = list(seg.get(u"uv") or [])
                if len(uv) < 2:
                    continue
                pts = [self._map_uv(p[0], p[1]) for p in uv]
                for i in range(len(pts) - 1):
                    d = _dist_point_seg(
                        px, py, pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1]
                    )
                    if d < best_d:
                        best_d = d
                        best = (bi, int(seg.get(u"idx") or 0))
        if best is None or best_d > _HIT_PX:
            return None
        return best

    def _select_hit(self, hit):
        if hit is None:
            self._selected = []
            self._update_meta()
            self._update_delete_enabled()
            self._redraw()
            return
        bi, si = hit
        self._selected = match_segments(self._session, bi, si)
        bars = list(self._session.get(u"bars") or [])
        n = 0
        if 0 <= bi < len(bars):
            n = int(bars[bi].get(u"n_segments") or 0)
        letter = segment_letter(si)
        if not is_end_segment(si, n):
            last = segment_letter(n - 1) if n > 0 else u"?"
            self._set_status(
                u"Tramo {0} · no válido (discontinuo). Elija un extremo (A o {1}).".format(
                    letter, last
                )
            )
        else:
            self._set_status(
                u"Tramo {0} en {1} barra(s) · listo para eliminar.".format(
                    letter, len(self._selected)
                )
            )
        self._update_meta()
        self._update_delete_enabled()
        self._redraw()

    def _canvas_pos(self, args):
        try:
            p = args.GetPosition(self._cnv)
            return float(p.X), float(p.Y)
        except Exception:
            return None, None

    def _on_canvas_down(self, sender, args):
        if self._panning:
            return
        try:
            if args.ChangedButton != MouseButton.Left:
                return
        except Exception:
            pass
        x, y = self._canvas_pos(args)
        if x is None:
            return
        self._select_hit(self._hit_test(x, y))

    def _on_canvas_move(self, sender, args):
        if self._panning and self._pan_last is not None:
            x, y = self._canvas_pos(args)
            if x is None:
                return
            self._pan_x += x - self._pan_last[0]
            self._pan_y += y - self._pan_last[1]
            self._pan_last = (x, y)
            self._redraw()
            return
        x, y = self._canvas_pos(args)
        if x is None:
            return
        hit = self._hit_test(x, y)
        if hit != self._hover:
            self._hover = hit
            self._redraw()

    def _on_canvas_leave(self, sender, args):
        if self._hover is not None:
            self._hover = None
            self._redraw()

    def _on_wheel(self, sender, args):
        try:
            delta = int(args.Delta)
        except Exception:
            return
        factor = _ZOOM_STEP if delta > 0 else 1.0 / _ZOOM_STEP
        nz = self._zoom * factor
        if nz < _ZOOM_MIN:
            nz = _ZOOM_MIN
        if nz > _ZOOM_MAX:
            nz = _ZOOM_MAX
        x, y = self._canvas_pos(args)
        if x is not None and self._zoom > 1e-9:
            # Zoom hacia el cursor
            k = nz / self._zoom
            self._pan_x = x - k * (x - self._pan_x)
            self._pan_y = y - k * (y - self._pan_y)
        self._zoom = nz
        self._redraw()

    def _on_mouse_down(self, sender, args):
        try:
            if args.ChangedButton == MouseButton.Middle:
                self._panning = True
                x, y = self._canvas_pos(args)
                self._pan_last = (x, y) if x is not None else None
                try:
                    self._cnv.CaptureMouse()
                except Exception:
                    pass
        except Exception:
            pass

    def _on_mouse_up(self, sender, args):
        try:
            if args.ChangedButton == MouseButton.Middle:
                self._panning = False
                self._pan_last = None
                try:
                    self._cnv.ReleaseMouseCapture()
                except Exception:
                    pass
        except Exception:
            pass

    def _on_host_size(self, sender, args):
        self._sync_canvas_size()
        self._redraw()

    def _load_rebars(self, uidoc, rebars):
        if not rebars:
            self._session = None
            self._selected = []
            self._hover = None
            self._zoom = 1.0
            self._pan_x = 0.0
            self._pan_y = 0.0
            self._update_meta()
            self._update_delete_enabled()
            self._redraw()
            self._set_status(u"Sin barras seleccionadas.")
            return
        view = resolve_active_model_view(uidoc)
        ok, err, session = build_elevation_session(uidoc.Document, rebars, view)
        if not ok:
            mostrar_aviso(self._uiapp, u"No se pudo dibujar el alzado.", err or u"")
            return
        self._session = session
        self._selected = []
        self._hover = None
        self._zoom = 1.0
        self._pan_x = 0.0
        self._pan_y = 0.0
        _clear_selection(uidoc)
        self._sync_canvas_size()
        self._update_meta()
        self._update_delete_enabled()
        self._redraw()
        skipped = list(session.get(u"skipped") or [])
        extra = u""
        if skipped:
            extra = u" ({0} omitida(s))".format(len(skipped))
        self._set_status(
            u"{0} barra(s) en alzado{1} · clic en un tramo.".format(
                len(session.get(u"bars") or []), extra
            )
        )

    def pick_in_revit(self, uiapp):
        uidoc = None
        try:
            uidoc = uiapp.ActiveUIDocument
        except Exception:
            uidoc = None
        hidden = False
        try:
            if self._win is not None:
                self._win.Hide()
                hidden = True
        except Exception:
            hidden = False
        try:
            rebars, err = pick_rebars(uidoc)
        except Exception as ex:
            rebars, err = [], _as_unicode(ex)
        finally:
            if hidden:
                try:
                    self._win.Show()
                    self._maximize()
                    self._win.Activate()
                except Exception:
                    pass
        if err:
            mostrar_aviso(uiapp, u"No se pudieron seleccionar barras.", err)
            return
        if not rebars:
            self._set_status(u"Selección cancelada.")
            return
        self._load_rebars(uidoc, rebars)

    def apply_in_revit(self, uiapp):
        uidoc = None
        try:
            uidoc = uiapp.ActiveUIDocument
        except Exception:
            uidoc = None
        if uidoc is None or not self._session or not self._selected:
            return
        bars = list(self._session.get(u"bars") or [])
        targets = []
        for bi, si in self._selected:
            if bi < 0 or bi >= len(bars):
                continue
            if not is_end_segment(si, int(bars[bi].get(u"n_segments") or 0)):
                continue
            targets.append((bars[bi].get(u"id"), int(si)))
        if not targets:
            mostrar_aviso(
                uiapp,
                u"No hay un tramo válido para eliminar.",
                u"Elija un extremo (A o el último) en el alzado.",
            )
            return
        view = resolve_active_model_view(uidoc)
        ok, msg, n_ok, errors, new_ids = apply_remove_segments(
            uidoc.Document, targets, view
        )
        if not ok:
            detalle = u"\n".join(_as_unicode(e) for e in (errors or [])[:8])
            mostrar_aviso(uiapp, msg or u"No se pudo eliminar el segmento.", detalle)
            return
        if errors:
            mostrar_aviso(uiapp, msg, u"\n".join(_as_unicode(e) for e in errors[:8]))
        try:
            self._win.Close()
        except Exception:
            _clear_singleton()

    def _on_manual(self, sender, args):
        _open_manual(self._uiapp)

    def _on_select(self, sender, args):
        if self._busy:
            return
        _set_hold(self)
        try:
            self._ext_pick.Raise()
        except Exception as ex:
            _clear_hold()
            mostrar_aviso(self._uiapp, u"No se pudo iniciar la selección.", _as_unicode(ex))

    def _on_delete(self, sender, args):
        if self._busy or not self._selection_valid():
            return
        _set_hold(self)
        try:
            self._ext_apply.Raise()
        except Exception as ex:
            _clear_hold()
            mostrar_aviso(self._uiapp, u"No se pudo aplicar.", _as_unicode(ex))

    def _on_close(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass

    def _on_closed(self, sender, args):
        _clear_singleton()

    def _prepare_on_revit_monitor(self):
        win = self._win
        if win is None:
            return
        hwnd = None
        try:
            hwnd = revit_main_hwnd(self._uiapp)
        except Exception:
            hwnd = None
        try:
            from System.Windows.Interop import WindowInteropHelper

            helper = WindowInteropHelper(win)
            if hwnd is not None:
                helper.Owner = hwnd
        except Exception:
            pass
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
        try:
            bind_maximize_wpf_on_revit_monitor(win, hwnd)
        except Exception:
            pass

    def _maximize(self):
        win = self._win
        if win is None:
            return
        try:
            win.SizeToContent = SizeToContent.Manual
        except Exception:
            pass
        try:
            win.WindowState = WindowState.Maximized
        except Exception:
            pass

    def show(self):
        self._prepare_on_revit_monitor()
        try:
            self._win.WindowState = WindowState.Maximized
        except Exception:
            pass
        self._win.Show()
        self._maximize()
        self._sync_canvas_size()
        self._redraw()
        try:
            self._win.Dispatcher.BeginInvoke(
                DispatcherPriority.ApplicationIdle,
                Action(self._sync_canvas_size),
            )
            self._win.Dispatcher.BeginInvoke(
                DispatcherPriority.ApplicationIdle,
                Action(self._redraw),
            )
        except Exception:
            pass


def show_eliminar_segmento_window(revit):
    """Pick múltiple primero; con barras elegidas abre la UI maximizada."""
    existing = _existing_ctrl()
    if existing is not None:
        if _activate_existing(existing, revit):
            return
    uidoc = None
    try:
        uidoc = revit.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        mostrar_aviso(revit, u"No hay documento activo.")
        return
    rebars, err = pick_rebars(uidoc)
    if err:
        mostrar_aviso(revit, u"No se pudieron seleccionar barras.", err)
        return
    if not rebars:
        return
    ctrl = EliminarSegmentoWindow(revit)
    ctrl._load_rebars(uidoc, rebars)
    if not ctrl._session:
        return
    _set_singleton(ctrl)
    ctrl.show()
