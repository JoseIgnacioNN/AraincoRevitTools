# -*- coding: utf-8 -*-
"""
Arainco: Fundación corrida (franja) — UI planta + franjas + colocación múltiple.

Carga Wall Foundations de la vista activa, dibuja franjas (polígonos) y coloca
armadura reutilizando el motor de ``enfierrado_wall_foundation``.

Revit 2024–2026 · IronPython (pyRevit).
"""

from __future__ import print_function

import math
import os
import sys
import time
import weakref

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

import System
from System import AppDomain, EventHandler
from System.Windows import (
    FontWeights,
    Point as WpfPoint,
    RoutedEventHandler,
    Thickness,
    Visibility,
    WindowState,
)
from System.Windows.Controls import Canvas as WpfCanvas
from System.Windows.Input import (
    Cursors,
    Key,
    Keyboard,
    KeyEventHandler,
    ModifierKeys,
    MouseButton,
    MouseButtonEventHandler,
    MouseEventHandler,
    MouseWheelEventHandler,
)
from System.Windows.Markup import XamlReader
from System.Windows.Media import Color, DoubleCollection, PointCollection, SolidColorBrush
from System.Windows.Shapes import Ellipse as WpfEllipse
from System.Windows.Shapes import Line as WpfLine
from System.Windows.Shapes import Polygon as WpfPolygon

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    Transaction,
    ViewPlan,
    WallFoundation,
    XYZ,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog

_scripts_dir = os.path.dirname(os.path.abspath(__file__))
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

from bimtools_ui_tokens import WINDOW_CHROME_TITLE
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_instruction_dialog import show_message_dialog
from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)
from fundacion_corrida_franja_geom import (
    assign_host_to_strip,
    build_host_preview_in_view,
    collect_wall_foundations_in_view,
    merge_strip_into_host_geo,
    point_in_polygon_mm,
    strip_axes_from_polygon_mm,
)
from canvas_sketch_osnap import (
    CanvasSketchOsnap,
    DRAW_POLY,
    DRAW_RECT,
    osnap_status_label,
)
from canvas_sketch_instrumentation import (
    CanvasSketchInstrument,
    dist_mm,
)
from barras_bordes_losa_gancho_empotramiento import (
    _build_bar_type_entries,
    _rebar_nominal_diameter_mm,
    element_id_to_int,
)
from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm
from area_rein_losa_panos import ensure_ccw, rect_from_two_points_mm, shoelace_area_m2
from conjunto_guid import (
    ARMADURA_UBICACION_INFERIOR,
    finalizar_armadura_conjunto_guid_ejecucion,
    iniciar_armadura_conjunto_guid_ejecucion,
    stamp_armadura_arainco,
    stamp_armadura_conjunto_guid,
    stamp_armadura_malla,
    stamp_armadura_nivel,
    stamp_armadura_posicion,
    stamp_armadura_ubicacion,
)

import enfierrado_wall_foundation as _ewf

_DIALOG_TITLE = u"Arainco: Fundación corrida (franja)"
_SINGLETON_KEY = u"Arainco.FundacionCorridaFranja.ActiveWindow"
_PLAN_PAD_FRAC = 0.08
_BRUSH_CACHE = {}
_COLOR_HOST = u"#5BC0DE"
_COLOR_STRIP = u"#4ade80"
_COLOR_STRIP_SEL = u"#fbbf24"
_COLOR_DRAFT = u"#fbbf24"
_COLOR_SNAP = u"#f472b6"
_SNAP_TOL_PX = 10.0  # apertura en pantalla (px), como Area Rein. Losa Sketch
# Misma convención que Area Rein. Losa Sketch (63_AreaReinLosaSketch)
_POLY_MIN_EDGE_MM = 25.0
_POLY_CLOSE_SNAP_MULT = 1.75
# Separación armadura (mm): valor numérico libre en el rail (sin forzar pasos)
_FRANJA_SEP_MM_MIN = 50
_FRANJA_SEP_MM_MAX = 2000
_FRANJA_SEP_MM_DEFAULT = 100
_FRANJA_SEP_MM_STEP = 10  # solo botones ▲/▼
_ARMADURA_EN_LAMINA_PARAM = u"Armadura_En Lamina"
# Cara inferior de zapata; U = luz menor (i); longitudinal = luz mayor (s)
_ARMADURA_POSICION_TRANS = u"i"
_ARMADURA_POSICION_LONG = u"s"


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _clamp_franja_sep_mm(raw_n, default_val=_FRANJA_SEP_MM_DEFAULT):
    try:
        n = int(round(float(raw_n)))
    except Exception:
        return int(default_val)
    return max(int(_FRANJA_SEP_MM_MIN), min(int(_FRANJA_SEP_MM_MAX), n))


def _normalize_franja_sep_tb(tb, default_val=_FRANJA_SEP_MM_DEFAULT):
    if tb is None:
        return
    try:
        s = _as_unicode(tb.Text).replace(u"mm", u"").strip()
        if not s:
            tb.Text = _as_unicode(int(default_val))
            return
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        tb.Text = _as_unicode(int(default_val))
        return
    tb.Text = _as_unicode(_clamp_franja_sep_mm(n, default_val))


def _read_franja_sep_tb(tb, default_val=_FRANJA_SEP_MM_DEFAULT):
    if tb is None:
        return int(default_val)
    try:
        s = _as_unicode(tb.Text).replace(u"mm", u"").strip()
        if not s:
            return int(default_val)
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        return int(default_val)
    return int(_clamp_franja_sep_mm(n, default_val))


def _sheet_number_from_view(view):
    """``Sheet Number`` de la vista (Armadura_En Lamina)."""
    if view is None:
        return u""
    try:
        p = view.LookupParameter(u"Sheet Number")
        if p is None:
            return u""
        try:
            if not p.HasValue:
                return u""
        except Exception:
            pass
        s = p.AsString()
        if s is not None and _as_unicode(s).strip():
            return _as_unicode(s).strip()
        vs = p.AsValueString()
        if vs is not None and _as_unicode(vs).strip():
            return _as_unicode(vs).strip()
    except Exception:
        pass
    return u""


def _nivel_nombre_desde_id(doc, eid):
    if doc is None or eid is None:
        return u""
    try:
        if eid == ElementId.InvalidElementId:
            return u""
    except Exception:
        pass
    try:
        el = doc.GetElement(eid)
    except Exception:
        return u""
    if el is None:
        return u""
    try:
        name = el.Name
        if name:
            return _as_unicode(name).strip()
    except Exception:
        pass
    return u""


def _nivel_nombre_wall_foundation(doc, wf):
    """Nivel asociado a la Wall Foundation (LevelId o base del muro host)."""
    if doc is None or wf is None:
        return u""
    try:
        lid = wf.LevelId
        name = _nivel_nombre_desde_id(doc, lid)
        if name:
            return name
    except Exception:
        pass
    try:
        wid = wf.WallId
        wall = doc.GetElement(wid) if wid is not None else None
    except Exception:
        wall = None
    if wall is not None:
        for bip_name in (
            u"WALL_BASE_CONSTRAINT",
            u"SCHEDULE_LEVEL_PARAM",
            u"LEVEL_PARAM",
        ):
            try:
                bip = getattr(BuiltInParameter, bip_name, None)
                if bip is None:
                    continue
                p = wall.get_Parameter(bip)
                if p is None or not p.HasValue:
                    continue
                name = _nivel_nombre_desde_id(doc, p.AsElementId())
                if name:
                    return name
            except Exception:
                continue
        for n in (u"Base Constraint", u"Restricción de base", u"Level", u"Nivel"):
            try:
                p = wall.LookupParameter(n)
                if p is None or not p.HasValue:
                    continue
                name = _nivel_nombre_desde_id(doc, p.AsElementId())
                if name:
                    return name
            except Exception:
                continue
    return u""


def _stamp_armadura_en_lamina(rebar, sheet_number):
    if rebar is None:
        return False
    try:
        valor = _as_unicode(sheet_number or u"").strip()
    except Exception:
        valor = u""
    try:
        p = rebar.LookupParameter(_ARMADURA_EN_LAMINA_PARAM)
        if p is None or p.IsReadOnly:
            return False
        p.Set(valor)
        return True
    except Exception:
        return False


def _stamp_rebars_franja(
    rebars,
    conjunto_guid,
    nivel_nombre,
    sheet_number,
    ubicacion,
    posicion,
):
    """Estampa parámetros Arainco en barras de una franja (U o longitudinales)."""
    if not rebars:
        return 0
    n = 0
    for rb in rebars:
        if rb is None:
            continue
        try:
            stamp_armadura_arainco(rb, yes=True)
        except Exception:
            pass
        try:
            stamp_armadura_malla(rb, yes=False)
        except Exception:
            pass
        try:
            if ubicacion:
                stamp_armadura_ubicacion(rb, ubicacion)
        except Exception:
            pass
        try:
            if posicion:
                stamp_armadura_posicion(rb, posicion)
        except Exception:
            pass
        try:
            if nivel_nombre:
                stamp_armadura_nivel(rb, nivel_nombre)
        except Exception:
            pass
        try:
            _stamp_armadura_en_lamina(rb, sheet_number)
        except Exception:
            pass
        try:
            if stamp_armadura_conjunto_guid(rb, conjunto_guid=conjunto_guid):
                n += 1
        except Exception:
            pass
    return n


def _brush(hex_color, alpha=255):
    h = (_as_unicode(hex_color) or u"#95B8CC").lstrip(u"#")
    if len(h) != 6:
        h = u"95B8CC"
    try:
        a = int(alpha)
    except Exception:
        a = 255
    key = (h, a)
    cached = _BRUSH_CACHE.get(key)
    if cached is not None:
        return cached
    brush = SolidColorBrush(
        Color.FromArgb(a, int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    )
    try:
        if brush.CanFreeze:
            brush.Freeze()
    except Exception:
        pass
    _BRUSH_CACHE[key] = brush
    return brush


def _mostrar_aviso(uiapp, instruction, content=u""):
    try:
        hwnd = revit_main_hwnd(uiapp)
    except Exception:
        hwnd = None
    try:
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
        TaskDialog.Show(_DIALOG_TITLE, _as_unicode(instruction))
    except Exception:
        pass


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401
    except Exception:
        return False
    return True


class _FundacionFranjaColocarProgress(object):
    """ProgressBar pyRevit (acento BIMTools); no-op si no está disponible."""

    def __init__(self, total, title_prefix=None):
        self._total = max(1, int(total or 1))
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or _DIALOG_TITLE

    def __enter__(self):
        if not _pbar_enabled():
            return self
        try:
            from pyrevit import forms as _pyrevit_forms

            self._pb = _pyrevit_forms.ProgressBar(
                title=self._title(0),
                cancellable=False,
            )
            try:
                self._pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(
                    Color.FromRgb(91, 192, 222),
                )
            except Exception:
                pass
            self._pb.__enter__()
            self._open = True
        except Exception:
            self._pb = None
            self._open = False
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._open and self._pb is not None:
            try:
                self._pb.__exit__(exc_type, exc_val, exc_tb)
            except Exception:
                pass
        self._open = False
        self._pb = None
        return False

    def _title(self, current):
        cur = max(0, int(current))
        if cur < 1:
            return u"{0} — Colocando 0/{1}…".format(
                self._title_prefix, int(self._total)
            )
        return u"{0} — Colocando {1}/{2}…".format(
            self._title_prefix, cur, int(self._total)
        )

    def update(self, current, label=None):
        if self._pb is None:
            return
        c = max(1, min(int(current), int(self._total)))
        base = self._title(c)
        if label:
            base = u"{0} ({1})".format(base, _as_unicode(label))
        try:
            if hasattr(self._pb, u"update_progress"):
                try:
                    self._pb.update_progress(c, max_value=self._total)
                except TypeError:
                    try:
                        self._pb.update_progress(c, max=self._total)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self._pb.title = base
        except Exception:
            pass


def _vista_es_planta(view):
    if view is None:
        return False
    try:
        return isinstance(view, ViewPlan)
    except Exception:
        return False


def _get_singleton():
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


_XAML = (
    u"""
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  xmlns:po="http://schemas.microsoft.com/winfx/2006/xaml/presentation/options"
  Title="__CHROME__"
  Height="900" Width="1280"
  MinHeight="720" MinWidth="1020"
  WindowStartupLocation="Manual"
  Background="#071018"
  FontFamily="Segoe UI"
  FontSize="12"
  ShowInTaskbar="False"
  ResizeMode="CanResize"
  UseLayoutRounding="True"
  SnapsToDevicePixels="True">
  <Window.Resources>
"""
    + BIMTOOLS_DARK_STYLES_XML
    + u"""
    <SolidColorBrush x:Key="FranjaAppBg" Color="#071018" po:Freeze="True"/>
    <SolidColorBrush x:Key="FranjaPanelBg" Color="#0a1620" po:Freeze="True"/>
    <SolidColorBrush x:Key="FranjaBorder" Color="#21465C" po:Freeze="True"/>
    <SolidColorBrush x:Key="FranjaFgHi" Color="#E8F4F8" po:Freeze="True"/>
    <SolidColorBrush x:Key="FranjaFgMid" Color="#95B8CC" po:Freeze="True"/>
    <SolidColorBrush x:Key="FranjaFgLo" Color="#64748b" po:Freeze="True"/>
    <SolidColorBrush x:Key="FranjaAccentSoft" Color="#7eb8d0" po:Freeze="True"/>
    <SolidColorBrush x:Key="FranjaInputBg" Color="#050E18" po:Freeze="True"/>
  </Window.Resources>
  <Border Background="{StaticResource FranjaAppBg}" BorderBrush="{StaticResource FranjaBorder}"
          BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Cabecera (mismo orden tipográfico que Armado vigas) -->
      <Grid Grid.Row="0" Margin="0,0,0,8">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="TxtTitle" Grid.Row="0"
                   Text="Arainco: Fundación corrida (franja)"
                   Foreground="{StaticResource FranjaFgHi}"
                   FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Grid.Row="1" Margin="0,6,0,0"
                   Foreground="{StaticResource FranjaFgMid}" FontSize="11"
                   TextWrapping="Wrap"
                   Text="Dibuje franjas sobre las zapatas de la planta · configure armadura por franja · Colocar cierra y modela."/>
      </Grid>

      <!-- Resumen compacto (hosts / franjas) -->
      <Border Grid.Row="1" Margin="0,0,0,8"
              Background="{StaticResource FranjaPanelBg}"
              BorderBrush="{StaticResource FranjaBorder}"
              BorderThickness="1" CornerRadius="4" Padding="10,7">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0">
            <TextBlock x:Name="TxtHostSummary"
                       Foreground="{StaticResource FranjaAccentSoft}"
                       FontSize="11" FontWeight="SemiBold"
                       Text="Wall Foundations en vista: —"/>
            <TextBlock x:Name="TxtStripSummary" Margin="0,3,0,0"
                       Foreground="{StaticResource FranjaFgLo}" FontSize="10"
                       Text="Franjas: 0"/>
          </StackPanel>
          <TextBlock Grid.Column="1" VerticalAlignment="Center"
                     Foreground="{StaticResource FranjaFgLo}" FontSize="9"
                     Text="Clic libre en franja = seleccionar · OSNAP = dibujar"
                     TextWrapping="Wrap" MaxWidth="220" TextAlignment="Right"/>
        </Grid>
      </Border>

      <TextBlock Grid.Row="2" x:Name="TxtHint"
                 Foreground="{StaticResource FranjaFgLo}" FontSize="10"
                 TextWrapping="Wrap" Margin="0,0,0,10"
                 Text="Rectángulo (2 clics; Shift = libre) · Polígono (n clics; cierre 1º / Enter) · Rueda = zoom · clic rueda = pan · Ctrl+0 = reset · Supr = borrar franja."/>

      <Grid Grid.Row="3">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="340"/>
        </Grid.ColumnDefinitions>

        <!-- Canvas planta -->
        <Border Grid.Column="0" Background="{StaticResource FranjaPanelBg}"
                BorderBrush="{StaticResource FranjaBorder}" BorderThickness="1"
                CornerRadius="4,0,0,4" Padding="0">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="{StaticResource FranjaPanelBg}"
                    BorderBrush="{StaticResource FranjaBorder}"
                    BorderThickness="0,0,0,1" Padding="10,7,10,6">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="TxtCanvasHeader" Grid.Column="0"
                           Foreground="{StaticResource FranjaFgLo}"
                           FontSize="10" FontWeight="SemiBold"
                           VerticalAlignment="Center"
                           Text="PLANTA · HUELLA EN VISTA"/>
                <Border Grid.Column="1" Background="{StaticResource FranjaInputBg}"
                        BorderBrush="{StaticResource FranjaBorder}" BorderThickness="1"
                        CornerRadius="4" Padding="3">
                  <StackPanel Orientation="Horizontal">
                    <Button x:Name="BtnDrawRect" Content="Rectángulo"
                            Style="{StaticResource BtnSelectOutline}"
                            MinWidth="92" Margin="0,0,3,0" Padding="10,4"
                            FontSize="11"
                            ToolTip="Franja con 2 clics (esquinas opuestas)"/>
                    <Button x:Name="BtnDrawPoly" Content="Polígono"
                            Style="{StaticResource BtnSelectOutline}"
                            MinWidth="84" Padding="10,4" FontSize="11"
                            ToolTip="Franja por vértices; cierre en 1º / Enter"/>
                  </StackPanel>
                </Border>
              </Grid>
            </Border>
            <TextBlock x:Name="TxtCanvasDebug" Grid.Row="1"
                       Visibility="Collapsed" Margin="10,4,10,0"
                       Foreground="{StaticResource FranjaFgLo}"
                       FontFamily="Consolas" FontSize="9"
                       TextWrapping="Wrap" MaxHeight="72"
                       Text="DEBUG canvas"/>
            <Border Grid.Row="2" Background="{StaticResource FranjaAppBg}"
                    Padding="8,6,8,8">
              <Border Background="{StaticResource FranjaAppBg}"
                      BorderBrush="{StaticResource FranjaBorder}"
                      BorderThickness="1" CornerRadius="4"
                      SnapsToDevicePixels="True">
                <Canvas x:Name="CvPlan" ClipToBounds="True" Focusable="True"
                        SnapsToDevicePixels="True"/>
              </Border>
            </Border>
          </Grid>
        </Border>

        <!-- Rail lateral -->
        <Border Grid.Column="1" Background="{StaticResource FranjaPanelBg}"
                BorderBrush="{StaticResource FranjaBorder}" BorderThickness="1"
                CornerRadius="0,4,4,0" Padding="10,10">
          <ScrollViewer VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Disabled"
                        CanContentScroll="False">
            <StackPanel>

              <Grid Margin="0,0,0,6">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" x:Name="TxtArmaduraTitle"
                           Text="Armadura"
                           Foreground="{StaticResource FranjaFgLo}"
                           FontSize="9" VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal"
                            VerticalAlignment="Center">
                  <TextBlock Text="Hormigón"
                             Foreground="{StaticResource FranjaFgLo}"
                             FontSize="9" VerticalAlignment="Center" Margin="0,0,6,0"/>
                  <ComboBox x:Name="CmbDosificacionHormigon"
                            Style="{StaticResource Combo}" MinWidth="72"
                            IsEditable="False" IsReadOnly="True">
                    <ComboBox.ItemContainerStyle>
                      <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
                    </ComboBox.ItemContainerStyle>
                  </ComboBox>
                </StackPanel>
              </Grid>

              <Border Background="{StaticResource FranjaPanelBg}"
                      BorderBrush="{StaticResource FranjaBorder}"
                      BorderThickness="1" CornerRadius="4" Padding="10">
                <StackPanel>
                  <TextBlock Text="Transversales (U)"
                             Foreground="{StaticResource FranjaFgMid}"
                             FontWeight="SemiBold" FontSize="11" Margin="0,0,0,6"/>
                  <TextBlock Text="Diámetro · separación (mm)"
                             Foreground="{StaticResource FranjaFgLo}"
                             FontSize="9" Margin="0,0,0,4"/>
                  <Grid Margin="0,0,0,14">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="108"/>
                    </Grid.ColumnDefinitions>
                    <ComboBox Grid.Column="0" x:Name="CmbTransDiam"
                              Style="{StaticResource Combo}"
                              IsEditable="False" IsReadOnly="True">
                      <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
                      </ComboBox.ItemContainerStyle>
                    </ComboBox>
                    <TextBlock Grid.Column="1" Text="@"
                               Foreground="{StaticResource FranjaFgMid}"
                               FontWeight="Bold" VerticalAlignment="Center"
                               Margin="6,0,6,0"/>
                    <Border Grid.Column="2" Height="28" CornerRadius="5"
                            Background="{StaticResource FranjaInputBg}"
                            BorderBrush="#1A3A4D" BorderThickness="1"
                            SnapsToDevicePixels="True">
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="18"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="TxtTransSep" Grid.Column="0"
                                 Style="{StaticResource CantSpinnerText}"
                                 Text="100" Padding="6,0,4,0"
                                 VerticalContentAlignment="Center"
                                 HorizontalAlignment="Stretch"
                                 ToolTip="Separación transversal (mm). Escriba un valor numérico."/>
                        <Border Grid.Column="1" Background="#11253D" BorderBrush="#1A3A4D"
                                BorderThickness="1,0,0,0" CornerRadius="0,5,5,0"
                                ClipToBounds="True">
                          <Grid>
                            <Grid.RowDefinitions>
                              <RowDefinition Height="*"/>
                              <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <RepeatButton x:Name="BtnTransSepUp" Grid.Row="0"
                                          Style="{StaticResource SpinRepeatBtn}" Content="▲"
                                          ToolTip="Más 10 mm"/>
                            <RepeatButton x:Name="BtnTransSepDown" Grid.Row="1"
                                          Style="{StaticResource SpinRepeatBtn}" Content="▼"
                                          ToolTip="Menos 10 mm"/>
                          </Grid>
                        </Border>
                      </Grid>
                    </Border>
                  </Grid>

                  <TextBlock Text="Longitudinales"
                             Foreground="{StaticResource FranjaFgMid}"
                             FontWeight="SemiBold" FontSize="11" Margin="0,0,0,6"/>
                  <TextBlock Text="Diámetro · separación (mm)"
                             Foreground="{StaticResource FranjaFgLo}"
                             FontSize="9" Margin="0,0,0,4"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="108"/>
                    </Grid.ColumnDefinitions>
                    <ComboBox Grid.Column="0" x:Name="CmbLongDiam"
                              Style="{StaticResource Combo}"
                              IsEditable="False" IsReadOnly="True">
                      <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
                      </ComboBox.ItemContainerStyle>
                    </ComboBox>
                    <TextBlock Grid.Column="1" Text="@"
                               Foreground="{StaticResource FranjaFgMid}"
                               FontWeight="Bold" VerticalAlignment="Center"
                               Margin="6,0,6,0"/>
                    <Border Grid.Column="2" Height="28" CornerRadius="5"
                            Background="{StaticResource FranjaInputBg}"
                            BorderBrush="#1A3A4D" BorderThickness="1"
                            SnapsToDevicePixels="True">
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="18"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="TxtLongSep" Grid.Column="0"
                                 Style="{StaticResource CantSpinnerText}"
                                 Text="100" Padding="6,0,4,0"
                                 VerticalContentAlignment="Center"
                                 HorizontalAlignment="Stretch"
                                 ToolTip="Separación longitudinal (mm). Escriba un valor numérico."/>
                        <Border Grid.Column="1" Background="#11253D" BorderBrush="#1A3A4D"
                                BorderThickness="1,0,0,0" CornerRadius="0,5,5,0"
                                ClipToBounds="True">
                          <Grid>
                            <Grid.RowDefinitions>
                              <RowDefinition Height="*"/>
                              <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <RepeatButton x:Name="BtnLongSepUp" Grid.Row="0"
                                          Style="{StaticResource SpinRepeatBtn}" Content="▲"
                                          ToolTip="Más 10 mm"/>
                            <RepeatButton x:Name="BtnLongSepDown" Grid.Row="1"
                                          Style="{StaticResource SpinRepeatBtn}" Content="▼"
                                          ToolTip="Menos 10 mm"/>
                          </Grid>
                        </Border>
                      </Grid>
                    </Border>
                  </Grid>

                  <Border x:Name="BorderTroceo" Visibility="Collapsed" Margin="0,12,0,0"
                          Background="{StaticResource FranjaAppBg}"
                          BorderBrush="{StaticResource FranjaBorder}" BorderThickness="1"
                          CornerRadius="4" Padding="8">
                    <StackPanel>
                      <TextBlock Text="Troceo (&gt; 12 m)"
                                 Foreground="{StaticResource FranjaFgMid}"
                                 FontSize="10" FontWeight="SemiBold" Margin="0,0,0,6"/>
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="8"/>
                          <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0">
                          <TextBlock Text="Largo máx. (mm)"
                                     Foreground="{StaticResource FranjaFgLo}" FontSize="9"
                                     Margin="0,0,0,3"/>
                          <TextBox x:Name="TxtMaxBarMm"
                                   Style="{StaticResource BimToolsTextBoxDark}"
                                   Text="12000" Height="26"/>
                        </StackPanel>
                        <StackPanel Grid.Column="2">
                          <TextBlock Text="Empalme (mm)"
                                     Foreground="{StaticResource FranjaFgLo}" FontSize="9"
                                     Margin="0,0,0,3"/>
                          <TextBox x:Name="TxtLapMm"
                                   Style="{StaticResource BimToolsTextBoxDark}"
                                   Text="600" Height="26"/>
                        </StackPanel>
                      </Grid>
                    </StackPanel>
                  </Border>
                </StackPanel>
              </Border>

            </StackPanel>
          </ScrollViewer>
        </Border>
      </Grid>

      <!-- Footer (Manual · estado · acciones) -->
      <Grid Grid.Row="4" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="BtnManual" Grid.Column="0" Content="Manual"
                Style="{StaticResource BtnSelectOutline}" MinWidth="96"
                Margin="0,0,12,0" Background="#2A5C3D"
                VerticalAlignment="Center"
                ToolTip="Abrir manual de usuario"/>
        <TextBlock x:Name="TxtStatus" Grid.Column="1" VerticalAlignment="Center"
                   Foreground="{StaticResource FranjaFgLo}" FontSize="10"
                   TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnCancelar" Content="Cancelar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnColocar" Content="Colocar armadura"
                  Style="{StaticResource BtnPrimary}" MinWidth="200"
                  ToolTip="Cierra la ventana y coloca armadura por franja"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
"""
).replace(u"__CHROME__", WINDOW_CHROME_TITLE)


class _ColocarFranjaHandler(IExternalEventHandler):
    def __init__(self):
        self._ctrl = None

    def GetName(self):
        return u"AraincoColocarFundacionCorridaFranja"

    def Execute(self, uiapp):
        ctrl = self._ctrl
        self._ctrl = None
        if ctrl is None:
            return
        try:
            ctrl._execute_colocar(uiapp)
        except Exception as ex:
            try:
                _mostrar_aviso(
                    uiapp,
                    u"Error al colocar armadura.",
                    content=_as_unicode(ex),
                )
            except Exception:
                pass
        finally:
            try:
                ctrl._colocar_pending = False
            except Exception:
                pass
            try:
                ctrl._dispose_col_event()
            except Exception:
                pass


class FundacionCorridaFranjaController(object):
    def __init__(self, uiapp, uidoc, doc, hosts, previews, source_view):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        self._hosts = list(hosts or [])
        self._previews = list(previews or [])  # parallel to hosts
        self._source_view = source_view
        self._view_frame = None
        for prev in self._previews:
            if prev and prev.get(u"view_frame"):
                self._view_frame = prev[u"view_frame"]
                break
        self._strips = []
        self._selected_strip_idx = None
        self._rail_loading = False
        self._draft_pts = []
        self._draw_mode = DRAW_POLY
        self._osnap = CanvasSketchOsnap()
        self._canvas_inst = CanvasSketchInstrument()
        self._hover_log_key = None
        self._hover_log_at = 0.0
        self._prev_ot_n = -1
        self._osnap_ring_count = 0
        self._hover_snap = None
        self._ui_btn_draw_rect = None
        self._ui_btn_draw_poly = None
        self._ui_txt_canvas_debug = None
        self._view_zoom = 1.0
        self._view_pan_x = 0.0
        self._view_pan_y = 0.0
        self._scene_base = None
        self._panning = False
        self._pan_last_x = 0.0
        self._pan_last_y = 0.0
        self._mouse_in_canvas = False
        self._entries = []
        self._win = None
        self._ui_cv = None
        self._col_handler = _ColocarFranjaHandler()
        self._col_event = ExternalEvent.Create(self._col_handler)
        self._colocar_pending = False

    def _collect_osnap_ring_polys(self):
        """Anillos OSNAP: hosts, recortes, franjas ya dibujadas."""
        rings = []
        for prev in self._previews or []:
            if not prev:
                continue
            polys = prev.get(u"polys")
            if not polys:
                p0 = prev.get(u"poly")
                polys = [p0] if p0 else []
            for poly in polys:
                if poly and len(poly) >= 3:
                    rings.append(list(poly))
            crop = prev.get(u"crop_uv")
            if crop is not None:
                try:
                    u0, u1, v0, v1 = crop
                    rings.append(
                        [
                            (float(u0), float(v0)),
                            (float(u1), float(v0)),
                            (float(u1), float(v1)),
                            (float(u0), float(v1)),
                        ]
                    )
                except Exception:
                    pass
        for strip in self._strips or []:
            poly = strip.get(u"poly") if strip else None
            if poly and len(poly) >= 3:
                rings.append(list(poly))
        self._osnap_ring_count = len(rings)
        return rings

    def _ensure_osnap_geometry(self):
        if not getattr(self._osnap, u"_dirty", True):
            return
        self._osnap.rebuild(
            self._collect_osnap_ring_polys(),
            draft_pts=self._draft_pts,
            track_origin=self._snap_track_origin(),
        )

    def _snap_track_origin(self):
        pts = self._draft_pts or []
        if not pts:
            return None
        if self._draw_mode == DRAW_POLY:
            return pts[-1]
        return pts[0]

    def _mark_osnap_dirty(self):
        self._osnap.mark_dirty()

    def _resolve_snap_mm(self, xmm, ymm, raw_only=False):
        if raw_only:
            return float(xmm), float(ymm), u"free"
        self._ensure_osnap_geometry()
        tol = self._snap_tol_mm()
        if tol <= 0.0:
            return float(xmm), float(ymm), None
        snapped, kind = self._osnap.resolve(
            (float(xmm), float(ymm)),
            self._snap_track_origin(),
            tol,
        )
        if snapped is None:
            return float(xmm), float(ymm), None
        return float(snapped[0]), float(snapped[1]), kind

    def _clear_pick_state(self, redraw=False):
        self._draft_pts = []
        self._hover_snap = None
        self._osnap.clear_tracks()
        self._mark_osnap_dirty()
        if redraw:
            self._refresh_summaries()
            self._redraw_plan()

    def _set_draw_mode(self, mode):
        mode = DRAW_POLY if mode == DRAW_POLY else DRAW_RECT
        prev = getattr(self, u"_draw_mode", DRAW_POLY)
        # Re-click del modo activo: no tocar el borrador (evita perder esquina A / vértices).
        if prev == mode:
            self._apply_draw_mode_visuals()
            try:
                self._canvas_dbg(
                    u"mode_set",
                    mode=mode,
                    action=u"noop_same",
                    draft_n=len(self._draft_pts or []),
                )
            except Exception:
                pass
            return
        cleared_n = len(self._draft_pts or [])
        self._clear_pick_state(redraw=False)
        self._draw_mode = mode
        self._apply_draw_mode_visuals()
        try:
            self._canvas_dbg(
                u"mode_set",
                mode=mode,
                prev=prev,
                action=u"switched",
                cleared_draft_n=cleared_n,
            )
        except Exception:
            pass
        if mode == DRAW_POLY:
            self._set_status(
                u"Modo Polígono: clic = vértice · cierre 1º/Enter · "
                u"OSNAP · Retroceso deshace · Esc cancela."
            )
        else:
            self._set_status(
                u"Modo Rectángulo: 2 clics en esquinas opuestas "
                u"(Shift+clic = punto libre)."
            )
        self._redraw_plan()

    def _apply_draw_mode_visuals(self):
        mode = getattr(self, u"_draw_mode", DRAW_POLY)
        for btn, active in (
            (self._ui_btn_draw_rect, mode == DRAW_RECT),
            (self._ui_btn_draw_poly, mode == DRAW_POLY),
        ):
            if btn is None:
                continue
            try:
                if active:
                    btn.Background = _brush(u"#164e63", 255)
                    btn.BorderBrush = _brush(u"#38bdf8", 255)
                    btn.Foreground = _brush(u"#E8F4F8", 255)
                    btn.FontWeight = FontWeights.SemiBold
                    btn.Opacity = 1.0
                else:
                    btn.Background = _brush(u"#0a1620", 255)
                    btn.BorderBrush = _brush(u"#21465C", 255)
                    btn.Foreground = _brush(u"#95B8CC", 255)
                    btn.FontWeight = FontWeights.Normal
                    btn.Opacity = 0.95
            except Exception:
                pass

    def _update_armadura_title(self):
        try:
            tb = self._win.FindName(u"TxtArmaduraTitle") if self._win else None
            if tb is None:
                return
            i = self._selected_strip_idx
            if i is not None and 0 <= i < len(self._strips):
                tb.Text = u"Armadura · Franja {0}".format(i + 1)
                try:
                    tb.Foreground = _brush(u"#7eb8d0", 255)
                    tb.FontWeight = FontWeights.SemiBold
                except Exception:
                    pass
            else:
                tb.Text = u"Armadura"
                try:
                    tb.Foreground = _brush(u"#64748b", 255)
                    tb.FontWeight = FontWeights.Normal
                except Exception:
                    pass
        except Exception:
            pass

    def _refresh_summaries(self):
        try:
            th = self._win.FindName(u"TxtHostSummary")
            if th is not None:
                th.Text = u"Wall Foundations en vista: {0}".format(len(self._hosts))
            ts = self._win.FindName(u"TxtStripSummary")
            if ts is not None:
                sel = self._selected_strip_idx
                draft_n = len(self._draft_pts or [])
                if sel is not None and 0 <= sel < len(self._strips):
                    ts.Text = u"Franjas: {0} · seleccionada: {1} · borrador: {2} vértice(s)".format(
                        len(self._strips), sel + 1, draft_n
                    )
                else:
                    ts.Text = u"Franjas: {0} · borrador: {1} vértice(s)".format(
                        len(self._strips), draft_n
                    )
        except Exception:
            pass
        # Troceo visible si alguna franja > 12 m
        try:
            from System.Windows import Visibility as Vis

            br = self._win.FindName(u"BorderTroceo")
            show = False
            for s in self._strips:
                if float(s.get(u"length_mm") or 0) > float(_ewf._MAX_STOCK_MM) + 0.01:
                    show = True
                    break
            if br is not None:
                br.Visibility = Vis.Visible if show else Vis.Collapsed
        except Exception:
            pass
        self._update_armadura_title()

    def _default_rail_config(self):
        return {
            u"trans_idx": 0,
            u"long_idx": 0,
            u"trans_sep_mm": int(_FRANJA_SEP_MM_DEFAULT),
            u"long_sep_mm": int(_FRANJA_SEP_MM_DEFAULT),
            u"grade": _ewf._DOSIFICACION_HORMIGON_DEFAULT,
            u"max_bar_mm": float(_ewf._MAX_STOCK_MM),
            u"lap_mm": float(_ewf._LAP_DEFAULT_MM),
        }

    def _read_rail_config(self):
        cfg = self._default_rail_config()
        win = self._win
        if win is None:
            return cfg
        try:
            cmb = win.FindName(u"CmbTransDiam")
            if cmb is not None:
                cfg[u"trans_idx"] = max(0, int(cmb.SelectedIndex))
        except Exception:
            pass
        try:
            cmb = win.FindName(u"CmbLongDiam")
            if cmb is not None:
                cfg[u"long_idx"] = max(0, int(cmb.SelectedIndex))
        except Exception:
            pass
        try:
            cfg[u"trans_sep_mm"] = int(
                _read_franja_sep_tb(win.FindName(u"TxtTransSep"))
            )
        except Exception:
            pass
        try:
            cfg[u"long_sep_mm"] = int(
                _read_franja_sep_tb(win.FindName(u"TxtLongSep"))
            )
        except Exception:
            pass
        try:
            cfg[u"grade"] = _ewf._read_dosificacion_hormigon(
                win.FindName(u"CmbDosificacionHormigon")
            )
        except Exception:
            pass
        try:
            cfg[u"max_bar_mm"] = float(
                _ewf._read_max_bar_tb(win.FindName(u"TxtMaxBarMm"))
            )
        except Exception:
            pass
        try:
            tlap = win.FindName(u"TxtLapMm")
            if tlap is not None:
                s = _as_unicode(tlap.Text).strip().replace(u",", u".")
                if s:
                    cfg[u"lap_mm"] = float(s)
        except Exception:
            pass
        return cfg

    def _apply_rail_config(self, cfg):
        if not cfg:
            cfg = self._default_rail_config()
        win = self._win
        if win is None:
            return
        self._rail_loading = True
        try:
            entries = self._entries or []
            n = len(entries)
            for name, key in (
                (u"CmbTransDiam", u"trans_idx"),
                (u"CmbLongDiam", u"long_idx"),
            ):
                cmb = win.FindName(name)
                if cmb is None:
                    continue
                try:
                    idx = int(cfg.get(key) or 0)
                except Exception:
                    idx = 0
                if n > 0:
                    idx = max(0, min(idx, n - 1))
                try:
                    cmb.SelectedIndex = idx
                except Exception:
                    pass
            try:
                tb = win.FindName(u"TxtTransSep")
                if tb is not None:
                    tb.Text = _as_unicode(
                        int(cfg.get(u"trans_sep_mm") or _FRANJA_SEP_MM_DEFAULT)
                    )
                    _normalize_franja_sep_tb(tb)
            except Exception:
                pass
            try:
                tb = win.FindName(u"TxtLongSep")
                if tb is not None:
                    tb.Text = _as_unicode(
                        int(cfg.get(u"long_sep_mm") or _FRANJA_SEP_MM_DEFAULT)
                    )
                    _normalize_franja_sep_tb(tb)
            except Exception:
                pass
            try:
                cmb_dos = win.FindName(u"CmbDosificacionHormigon")
                if cmb_dos is not None:
                    grade = _as_unicode(cfg.get(u"grade") or _ewf._DOSIFICACION_HORMIGON_DEFAULT)
                    opts = list(_ewf._DOSIFICACION_HORMIGON_OPCIONES or [])
                    if grade in opts:
                        cmb_dos.SelectedIndex = opts.index(grade)
                    else:
                        cmb_dos.SelectedIndex = 0
            except Exception:
                pass
            try:
                tmax = win.FindName(u"TxtMaxBarMm")
                if tmax is not None:
                    tmax.Text = _as_unicode(int(round(float(cfg.get(u"max_bar_mm") or _ewf._MAX_STOCK_MM))))
                    _ewf._normalize_max_bar_tb(tmax)
            except Exception:
                pass
            try:
                tlap = win.FindName(u"TxtLapMm")
                if tlap is not None:
                    tlap.Text = _as_unicode(int(round(float(cfg.get(u"lap_mm") or _ewf._LAP_DEFAULT_MM))))
                    _ewf._normalize_lap_tb(tlap)
            except Exception:
                pass
        finally:
            self._rail_loading = False

    def _sync_rail_to_selected_strip(self):
        if self._rail_loading:
            return
        i = self._selected_strip_idx
        if i is None or i < 0 or i >= len(self._strips):
            return
        try:
            self._strips[i][u"config"] = self._read_rail_config()
        except Exception:
            pass

    def _clear_strip_selection(self, status=None):
        self._selected_strip_idx = None
        self._update_armadura_title()
        self._redraw_plan()
        if status:
            self._set_status(status)

    def _select_strip(self, idx):
        if idx is None or idx < 0 or idx >= len(self._strips):
            self._clear_strip_selection()
            return
        self._selected_strip_idx = int(idx)
        strip = self._strips[self._selected_strip_idx]
        cfg = strip.get(u"config")
        if not cfg:
            cfg = self._read_rail_config()
            strip[u"config"] = cfg
        self._apply_rail_config(cfg)
        self._update_armadura_title()
        self._refresh_summaries()
        self._redraw_plan()
        self._set_status(u"Franja {0} seleccionada.".format(self._selected_strip_idx + 1))

    def _hit_strip_at_mm(self, pt):
        """Índice de franja bajo el punto; si solapan, la de menor área (luego la última)."""
        if not pt or not self._strips:
            return None
        try:
            px = float(pt[0])
            py = float(pt[1])
        except Exception:
            return None
        candidates = []
        for i, s in enumerate(self._strips):
            poly = s.get(u"poly")
            if not poly or len(poly) < 3:
                continue
            try:
                if not point_in_polygon_mm(px, py, poly):
                    continue
            except Exception:
                continue
            try:
                area = abs(float(shoelace_area_m2(poly)))
            except Exception:
                try:
                    area = (
                        float(s.get(u"length_mm") or 0)
                        * float(s.get(u"width_mm") or 0)
                        / 1.0e6
                    )
                except Exception:
                    area = 1.0e9
            # área ascendente; -i hace preferir índice mayor si empatan
            candidates.append((area, -i, i))
        if not candidates:
            return None
        candidates.sort()
        return candidates[0][2]

    def _resolver_bar_type_from_idx(self, document, entries, idx):
        entries = entries or []
        try:
            i = int(idx)
        except Exception:
            i = 0
        if 0 <= i < len(entries):
            bt, lbl = entries[i]
            if bt is not None:
                return bt, None
            try:
                mm = _ewf._parse_diameter_mm_from_bar_combo_label(lbl)
            except Exception:
                mm = None
            if mm is not None and document is not None:
                try:
                    from enfierrado_shaft_hashtag import resolver_bar_type_por_diametro_mm

                    bt2, _, _ = resolver_bar_type_por_diametro_mm(document, float(mm))
                    if bt2 is not None:
                        return bt2, None
                except Exception:
                    pass
        return None, u"No se pudo resolver RebarBarType."

    def _on_rail_changed(self, sender=None, args=None):
        if self._rail_loading:
            return
        self._sync_rail_to_selected_strip()

    def _on_long_diam_changed(self, sender=None, args=None):
        if self._rail_loading:
            return
        self._sync_lap()
        self._sync_rail_to_selected_strip()

    def _on_sep_lost_focus(self, sender, args, tbx=None):
        if tbx is not None:
            try:
                _normalize_franja_sep_tb(tbx)
            except Exception:
                pass
        self._on_rail_changed()

    def _step_franja_sep(self, tb, delta):
        if tb is None:
            return
        try:
            v = _read_franja_sep_tb(tb)
        except Exception:
            v = int(_FRANJA_SEP_MM_DEFAULT)
        try:
            v = int(v) + int(delta)
        except Exception:
            v = int(_FRANJA_SEP_MM_DEFAULT)
        tb.Text = _as_unicode(_clamp_franja_sep_mm(v))
        self._on_rail_changed()

    def _keyboard_focus_in_textbox(self):
        try:
            from System.Windows.Controls import TextBox as _WpfTb

            fe = Keyboard.FocusedElement
            return fe is not None and isinstance(fe, _WpfTb)
        except Exception:
            return False

    def _on_max_lost_focus(self, sender, args):
        try:
            _ewf._normalize_max_bar_tb(self._win.FindName(u"TxtMaxBarMm"))
        except Exception:
            pass
        self._on_rail_changed()

    def _on_lap_lost_focus(self, sender, args):
        try:
            _ewf._normalize_lap_tb(self._win.FindName(u"TxtLapMm"))
        except Exception:
            pass
        self._on_rail_changed()

    def _commit_strip_from_pts(self, pts, snap_lbl=u""):
        if not pts or len(pts) < 3:
            try:
                self._canvas_dbg(
                    u"strip_commit_fail",
                    reason=u"geom",
                    pts_n=0 if not pts else len(pts),
                    strips=len(self._strips or []),
                )
            except Exception:
                pass
            self._set_status(u"No se creó franja: geometría insuficiente.")
            return False
        area = 0.0
        try:
            area = float(shoelace_area_m2(pts))
        except Exception:
            area = 0.0
        if area < 1e-6:
            try:
                self._canvas_dbg(
                    u"strip_commit_fail",
                    reason=u"area",
                    pts_n=len(pts),
                    area=u"{0:.6f}".format(area),
                    strips=len(self._strips or []),
                )
            except Exception:
                pass
            self._set_status(u"No se creó franja: área demasiado pequeña.")
            return False
        strip = strip_axes_from_polygon_mm(list(pts))
        if strip is None:
            try:
                self._canvas_dbg(
                    u"strip_commit_fail",
                    reason=u"axes",
                    pts_n=len(pts),
                    area=u"{0:.6f}".format(area),
                    strips=len(self._strips or []),
                )
            except Exception:
                pass
            self._set_status(u"No se pudo interpretar la franja.")
            return False
        strip[u"config"] = self._read_rail_config()
        self._strips.append(strip)
        self._selected_strip_idx = len(self._strips) - 1
        self._clear_pick_state(redraw=False)
        self._update_armadura_title()
        self._refresh_summaries()
        self._redraw_plan()
        try:
            self._canvas_dbg(
                u"strip_commit",
                strips=len(self._strips),
                length_mm=u"{0:.1f}".format(float(strip[u"length_mm"])),
                width_mm=u"{0:.1f}".format(float(strip[u"width_mm"])),
                pts_n=len(pts),
            )
        except Exception:
            pass
        lbl = snap_lbl or u""
        if lbl and not lbl.startswith(u" ·"):
            lbl = u" ·" + lbl
        self._set_status(
            u"Franja {0}: {1:.0f} × {2:.0f} mm{3}.".format(
                len(self._strips),
                float(strip[u"length_mm"]),
                float(strip[u"width_mm"]),
                lbl,
            )
        )
        return True

    def _draw_osnap_guides(self, cv, to_px):
        guides = getattr(self._osnap, u"guides", None)
        if cv is None or not guides:
            return
        for g in guides:
            try:
                p0, p1 = g[0], g[1]
                a = to_px(p0[0], p0[1])
                b = to_px(p1[0], p1[1])
            except Exception:
                continue
            ln = WpfLine()
            ln.X1, ln.Y1, ln.X2, ln.Y2 = float(a[0]), float(a[1]), float(b[0]), float(b[1])
            ln.Stroke = _brush(u"#38bdf8", 150)
            ln.StrokeThickness = 0.85
            try:
                dashes = DoubleCollection()
                dashes.Add(6)
                dashes.Add(4)
                ln.StrokeDashArray = dashes
            except Exception:
                pass
            try:
                cv.Children.Add(ln)
            except Exception:
                pass
            try:
                ln.IsHitTestVisible = False
            except Exception:
                pass

    def _set_status(self, text):
        try:
            tb = self._win.FindName(u"TxtStatus") if self._win else None
            if tb is not None:
                tb.Text = _as_unicode(text)
        except Exception:
            pass

    def _cargar_combos(self):
        entries, err = _build_bar_type_entries(self._doc)
        if err:
            _mostrar_aviso(self._uiapp, err)
            entries = entries or []
        self._entries = entries
        for name in (u"CmbTransDiam", u"CmbLongDiam"):
            cmb = self._win.FindName(name)
            if cmb is None:
                continue
            try:
                cmb.Items.Clear()
            except Exception:
                pass
            for _bt, lbl in entries:
                try:
                    cmb.Items.Add(lbl)
                except Exception:
                    pass
            try:
                cmb.SelectedIndex = 0
            except Exception:
                pass
        cmb_dos = self._win.FindName(u"CmbDosificacionHormigon")
        if cmb_dos is not None:
            try:
                cmb_dos.Items.Clear()
                for lab in _ewf._DOSIFICACION_HORMIGON_OPCIONES:
                    cmb_dos.Items.Add(lab)
                cmb_dos.SelectedIndex = 0
            except Exception:
                pass
        _normalize_franja_sep_tb(self._win.FindName(u"TxtTransSep"))
        _normalize_franja_sep_tb(self._win.FindName(u"TxtLongSep"))

    def _canvas_size(self, cv):
        w = h = 40.0
        try:
            w = float(cv.ActualWidth or 0)
            h = float(cv.ActualHeight or 0)
        except Exception:
            pass
        return max(40.0, w), max(40.0, h)

    def _plan_bbox_mm(self):
        xs = []
        ys = []
        for prev in self._previews:
            if not prev:
                continue
            for poly in prev.get(u"polys") or [prev.get(u"poly")]:
                if not poly:
                    continue
                for x, y in poly:
                    xs.append(float(x))
                    ys.append(float(y))
            crop = prev.get(u"crop_uv")
            if crop is not None:
                try:
                    u0, u1, v0, v1 = crop
                    xs.extend([float(u0), float(u1)])
                    ys.extend([float(v0), float(v1)])
                except Exception:
                    pass
        for s in self._strips:
            for x, y in s.get(u"poly") or []:
                xs.append(float(x))
                ys.append(float(y))
        for x, y in self._draft_pts:
            xs.append(float(x))
            ys.append(float(y))
        if not xs:
            return 0.0, 1000.0, 0.0, 1000.0
        return min(xs), max(xs), min(ys), max(ys)

    def _compute_plan_view_transform(self):
        cv = self._ui_cv
        if cv is None:
            return None
        cw, ch = self._canvas_size(cv)
        min_x, max_x, min_y, max_y = self._plan_bbox_mm()
        span_x = max(1.0, max_x - min_x)
        span_y = max(1.0, max_y - min_y)
        pad = _PLAN_PAD_FRAC
        fit = min(
            (cw * (1.0 - 2.0 * pad)) / span_x,
            (ch * (1.0 - 2.0 * pad)) / span_y,
        )
        fit = max(1e-6, fit)
        scale = fit * max(0.05, float(self._view_zoom))
        cx_mm = 0.5 * (min_x + max_x) + float(self._view_pan_x)
        cy_mm = 0.5 * (min_y + max_y) + float(self._view_pan_y)
        ox = cw / 2.0 - (cx_mm - min_x) * scale
        oy = ch / 2.0 - (max_y - cy_mm) * scale
        return {
            u"min_x": min_x,
            u"max_x": max_x,
            u"min_y": min_y,
            u"max_y": max_y,
            u"ox": ox,
            u"oy": oy,
            u"scale": scale,
            u"fit": fit,
        }

    def _ensure_scene_transform(self):
        sb = self._scene_base
        if sb is not None:
            try:
                if float(sb.get(u"scale") or 0) > 1e-12:
                    return sb
            except Exception:
                pass
        sb = self._compute_plan_view_transform()
        if sb is not None:
            self._scene_base = sb
        return sb

    def _canvas_mouse_active(self):
        try:
            return self._ui_cv is not None and bool(self._ui_cv.IsMouseOver)
        except Exception:
            return bool(getattr(self, u"_mouse_in_canvas", False))

    def _redraw_plan(self):
        cv = self._ui_cv
        if cv is None:
            return
        try:
            cv.Children.Clear()
        except Exception:
            return
        sb = self._compute_plan_view_transform()
        if sb is None:
            return
        self._scene_base = sb
        min_x = float(sb[u"min_x"])
        max_y = float(sb[u"max_y"])
        ox = float(sb[u"ox"])
        oy = float(sb[u"oy"])
        scale = float(sb[u"scale"])

        def to_px(xmm, ymm):
            return (
                ox + (float(xmm) - min_x) * scale,
                oy + (max_y - float(ymm)) * scale,
            )

        def add_poly(pts, stroke, fill_a, thick=1.5):
            if not pts or len(pts) < 2:
                return
            wp = WpfPolygon()
            pc = PointCollection()
            for xmm, ymm in pts:
                px, py = to_px(xmm, ymm)
                pc.Add(WpfPoint(px, py))
            wp.Points = pc
            wp.Stroke = _brush(stroke)
            wp.StrokeThickness = thick
            if fill_a > 0:
                wp.Fill = _brush(stroke, fill_a)
            try:
                wp.IsHitTestVisible = False
            except Exception:
                pass
            try:
                cv.Children.Add(wp)
            except Exception:
                pass

        crop_drawn = False
        for prev in self._previews:
            if not prev:
                continue
            for poly in prev.get(u"polys") or [prev.get(u"poly")]:
                if poly:
                    add_poly(poly, _COLOR_HOST, 100, 1.4)
            if (not crop_drawn) and prev.get(u"crop_uv") is not None:
                try:
                    u0, u1, v0, v1 = prev[u"crop_uv"]
                    add_poly(
                        [(u0, v0), (u1, v0), (u1, v1), (u0, v1)],
                        u"#64748b",
                        0,
                        1.0,
                    )
                    crop_drawn = True
                except Exception:
                    pass

        for i, s in enumerate(self._strips):
            poly = s.get(u"poly")
            selected = self._selected_strip_idx is not None and i == self._selected_strip_idx
            stroke = _COLOR_STRIP_SEL if selected else _COLOR_STRIP
            fill_a = 120 if selected else 80
            thick = 2.8 if selected else 1.8
            if poly:
                add_poly(poly, stroke, fill_a, thick)
            p0 = s.get(u"p0_mm")
            p1 = s.get(u"p1_mm")
            if p0 and p1:
                ln = WpfLine()
                x0, y0 = to_px(p0[0], p0[1])
                x1, y1 = to_px(p1[0], p1[1])
                ln.X1, ln.Y1, ln.X2, ln.Y2 = x0, y0, x1, y1
                ln.Stroke = _brush(stroke)
                ln.StrokeThickness = 3.0 if selected else 2.2
                try:
                    ln.IsHitTestVisible = False
                    cv.Children.Add(ln)
                except Exception:
                    pass

        # Guías OSNAP (perpendicular, proyección, tracking)
        self._draw_osnap_guides(cv, to_px)

        # Borrador — polígono multipunto o preview rectángulo
        pick_pts = list(self._draft_pts or [])
        hover = self._hover_snap
        draw_mode = getattr(self, u"_draw_mode", DRAW_POLY)
        if pick_pts and draw_mode == DRAW_RECT and hover is not None:
            try:
                preview = rect_from_two_points_mm(pick_pts[0], (hover[0], hover[1]))
                if preview and len(preview) >= 4:
                    add_poly(preview, _COLOR_DRAFT, 45, 0.9)
            except Exception:
                pass
        elif pick_pts:
            try:
                chain = list(pick_pts)
                hx = hy = None
                if hover is not None:
                    try:
                        hx, hy = float(hover[0]), float(hover[1])
                        chain.append((hx, hy))
                    except Exception:
                        hx = hy = None
                nchain = len(chain)
                for i in range(max(0, nchain - 1)):
                    x0, y0 = to_px(chain[i][0], chain[i][1])
                    x1, y1 = to_px(chain[i + 1][0], chain[i + 1][1])
                    is_rubber = hx is not None and i == nchain - 2
                    ln = WpfLine()
                    ln.X1, ln.Y1, ln.X2, ln.Y2 = x0, y0, x1, y1
                    ln.Stroke = _brush(
                        _COLOR_DRAFT, 220 if not is_rubber else 160
                    )
                    ln.StrokeThickness = 1.1 if not is_rubber else 0.9
                    if is_rubber:
                        try:
                            dashes = DoubleCollection()
                            dashes.Add(5)
                            dashes.Add(3)
                            ln.StrokeDashArray = dashes
                        except Exception:
                            pass
                    try:
                        ln.IsHitTestVisible = False
                        cv.Children.Add(ln)
                    except Exception:
                        pass
                if len(pick_pts) >= 2 and hx is not None:
                    x0, y0 = to_px(hx, hy)
                    x1, y1 = to_px(pick_pts[0][0], pick_pts[0][1])
                    near_close = self._point_near_first_vertex((hx, hy))
                    ln = WpfLine()
                    ln.X1, ln.Y1, ln.X2, ln.Y2 = x0, y0, x1, y1
                    ln.Stroke = _brush(
                        u"#4ade80" if near_close else _COLOR_DRAFT,
                        200 if near_close else 100,
                    )
                    ln.StrokeThickness = 1.2 if near_close else 0.75
                    try:
                        dashes = DoubleCollection()
                        dashes.Add(4)
                        dashes.Add(3)
                        ln.StrokeDashArray = dashes
                    except Exception:
                        pass
                    try:
                        ln.IsHitTestVisible = False
                        cv.Children.Add(ln)
                    except Exception:
                        pass
                for i, (xmm, ymm) in enumerate(pick_pts):
                    px, py = to_px(xmm, ymm)
                    r = 4.5 if i == 0 else 3.5
                    el = WpfEllipse()
                    el.Width = el.Height = r * 2.0
                    if i == 0:
                        el.Fill = _brush(u"#4ade80", 210)
                    else:
                        el.Fill = _brush(_COLOR_DRAFT, 200)
                    el.Stroke = _brush(u"#E8F4F8")
                    el.StrokeThickness = 0.8
                    try:
                        el.IsHitTestVisible = False
                        WpfCanvas.SetLeft(el, px - r)
                        WpfCanvas.SetTop(el, py - r)
                        cv.Children.Add(el)
                    except Exception:
                        pass
            except Exception:
                pass

        # Marcador de snap (hover) — solo con snap activo o borrador en curso
        hs = self._hover_snap
        if hs is not None and (pick_pts or hs[2] is not None):
            try:
                hx, hy = float(hs[0]), float(hs[1])
                px, py = to_px(hx, hy)
                arm = 8.0
                ln1 = WpfLine()
                ln1.X1, ln1.Y1, ln1.X2, ln1.Y2 = px - arm, py, px + arm, py
                ln1.Stroke = _brush(_COLOR_SNAP)
                ln1.StrokeThickness = 1.6
                ln2 = WpfLine()
                ln2.X1, ln2.Y1, ln2.X2, ln2.Y2 = px, py - arm, px, py + arm
                ln2.Stroke = _brush(_COLOR_SNAP)
                ln2.StrokeThickness = 1.6
                ring = WpfEllipse()
                ring.Width = ring.Height = 12
                ring.Stroke = _brush(_COLOR_SNAP)
                ring.StrokeThickness = 1.8
                ring.Fill = _brush(_COLOR_SNAP, 60)
                WpfCanvas.SetLeft(ring, px - 6)
                WpfCanvas.SetTop(ring, py - 6)
                for el in (ln1, ln2, ring):
                    try:
                        el.IsHitTestVisible = False
                    except Exception:
                        pass
                cv.Children.Add(ln1)
                cv.Children.Add(ln2)
                cv.Children.Add(ring)
            except Exception:
                pass

        try:
            hdr = self._win.FindName(u"TxtCanvasHeader")
            if hdr is not None:
                hdr.Text = (
                    u"PLANTA · {0} zapata(s) · {1} franja(s)"
                ).format(len(self._hosts), len(self._strips))
        except Exception:
            pass

    def _px_to_mm(self, px, py):
        sb = self._ensure_scene_transform()
        if sb is None:
            return None
        scale = float(sb.get(u"scale") or 0)
        if scale < 1e-12:
            return None
        min_x = float(sb[u"min_x"])
        max_y = float(sb[u"max_y"])
        ox = float(sb[u"ox"])
        oy = float(sb[u"oy"])
        xmm = min_x + (float(px) - ox) / scale
        ymm = max_y - (float(py) - oy) / scale
        return (xmm, ymm)

    def _snap_tol_mm(self):
        sb = self._ensure_scene_transform()
        if sb is None:
            return 0.0
        scale = float(sb.get(u"scale") or 0)
        if scale < 1e-12:
            return 0.0
        return float(_SNAP_TOL_PX) / scale

    def _canvas_dbg(self, event, **fields):
        try:
            self._canvas_inst.event(event, **fields)
            self._refresh_canvas_debug_ui()
        except Exception:
            pass

    def _refresh_canvas_debug_ui(self):
        tb = self._ui_txt_canvas_debug
        if tb is None:
            return
        try:
            from System.Windows import Visibility

            if self._canvas_inst.is_ui_enabled():
                tb.Visibility = Visibility.Visible
                tb.Text = self._canvas_inst.ui_text()
            else:
                tb.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def _canvas_view_snapshot(self):
        cv = self._ui_cv
        sb = self._scene_base
        snap = {}
        try:
            snap = self._osnap.stats()
        except Exception:
            pass
        fields = {
            u"zoom": u"{0:.3f}".format(float(self._view_zoom or 1.0)),
            u"pan_x": u"{0:.1f}".format(float(self._view_pan_x or 0.0)),
            u"pan_y": u"{0:.1f}".format(float(self._view_pan_y or 0.0)),
            u"tol_mm": u"{0:.2f}".format(float(self._snap_tol_mm())),
            u"mode": getattr(self, u"_draw_mode", DRAW_POLY),
            u"draft_n": len(self._draft_pts or []),
            u"strips": len(self._strips or []),
        }
        if sb is not None:
            try:
                fields[u"scale"] = u"{0:.6f}".format(float(sb.get(u"scale") or 0.0))
                fields[u"fit"] = u"{0:.6f}".format(float(sb.get(u"fit") or 0.0))
            except Exception:
                pass
        if cv is not None:
            try:
                fields[u"cv_w"] = u"{0:.0f}".format(float(cv.ActualWidth or 0.0))
                fields[u"cv_h"] = u"{0:.0f}".format(float(cv.ActualHeight or 0.0))
                fields[u"mouse_over"] = bool(cv.IsMouseOver)
                fields[u"focused"] = bool(cv.IsFocused)
                fields[u"captured"] = bool(cv.IsMouseCaptured)
            except Exception:
                pass
        for key, val in snap.items():
            fields[u"osnap_" + _as_unicode(key)] = val
        fields[u"osnap_rings"] = int(getattr(self, u"_osnap_ring_count", 0) or 0)
        return fields

    def _canvas_click_diag(
        self, pos, pt, sx, sy, kind, outcome, last_pt=None, **extra
    ):
        fields = self._canvas_view_snapshot()
        try:
            fields[u"px"] = u"{0:.1f},{1:.1f}".format(float(pos.X), float(pos.Y))
        except Exception:
            pass
        if pt is not None:
            fields[u"raw_mm"] = u"{0:.1f},{1:.1f}".format(float(pt[0]), float(pt[1]))
        fields[u"snap_mm"] = u"{0:.1f},{1:.1f}".format(float(sx), float(sy))
        fields[u"snap_kind"] = kind or u"none"
        fields[u"outcome"] = outcome
        if pt is not None:
            dmm = dist_mm(pt, (sx, sy))
            if dmm is not None:
                fields[u"snap_delta_mm"] = u"{0:.2f}".format(float(dmm))
        pick_pts = self._draft_pts or []
        ref_last = last_pt
        if ref_last is None and pick_pts:
            ref_last = pick_pts[-1]
        if ref_last is not None:
            try:
                fields[u"dist_last_mm"] = u"{0:.2f}".format(
                    float(dist_mm((sx, sy), ref_last) or 0.0)
                )
            except Exception:
                pass
        if pick_pts:
            try:
                fx, fy = float(pick_pts[0][0]), float(pick_pts[0][1])
                d_first = dist_mm((sx, sy), (fx, fy))
                if d_first is not None:
                    fields[u"dist_first_mm"] = u"{0:.2f}".format(float(d_first))
                fields[u"close_thr_mm"] = u"{0:.2f}".format(
                    float(self._poly_close_thresh_mm())
                )
            except Exception:
                pass
        for key, val in extra.items():
            fields[key] = val
        self._canvas_dbg(u"click", **fields)

    def _toggle_canvas_debug(self):
        on = self._canvas_inst.toggle_ui()
        self._refresh_canvas_debug_ui()
        state = u"ON" if on else u"OFF"
        path = self._canvas_inst.log_path
        self._canvas_dbg(
            u"debug_toggle",
            ui=state,
            log=path,
        )
        self._set_status(
            u"Debug canvas {0}. Log: {1} · Ctrl+Shift+L abre el archivo.".format(
                state, path
            )
        )
        return on

    def _open_canvas_debug_log(self):
        path = self._canvas_inst.log_path
        try:
            if not os.path.isfile(path):
                self._canvas_dbg(u"log_open", created=1, path=path)
            os.startfile(path)
        except Exception as ex:
            _mostrar_aviso(
                self._uiapp,
                u"No se pudo abrir el log de canvas.",
                _as_unicode(ex),
            )

    def _snap_kind_label(self, kind):
        return osnap_status_label(kind).strip()

    def _poly_close_thresh_mm(self):
        thr = self._snap_tol_mm()
        if thr <= 0:
            thr = 50.0
        return max(float(thr) * float(_POLY_CLOSE_SNAP_MULT), 15.0)

    def _point_near_first_vertex(self, pt):
        pts = self._draft_pts or []
        if not pts or pt is None:
            return False
        try:
            dx = float(pt[0]) - float(pts[0][0])
            dy = float(pt[1]) - float(pts[0][1])
        except Exception:
            return False
        r = self._poly_close_thresh_mm()
        return (dx * dx + dy * dy) <= (r * r)

    def _normalize_poly_pts_mm(self, pts):
        """Limpia vértices duplicados y asegura orientación CCW (Losa Sketch)."""
        if not pts or len(pts) < 3:
            return None
        cleaned = []
        min_e = float(_POLY_MIN_EDGE_MM)
        min_e2 = min_e * min_e
        for p in pts:
            try:
                x, y = float(p[0]), float(p[1])
            except Exception:
                continue
            if cleaned:
                dx = x - cleaned[-1][0]
                dy = y - cleaned[-1][1]
                if dx * dx + dy * dy < min_e2:
                    continue
            cleaned.append((x, y))
        if len(cleaned) >= 3:
            dx = cleaned[0][0] - cleaned[-1][0]
            dy = cleaned[0][1] - cleaned[-1][1]
            if dx * dx + dy * dy < min_e2:
                cleaned = cleaned[:-1]
        if len(cleaned) < 3:
            return None
        try:
            return ensure_ccw(cleaned)
        except Exception:
            return cleaned

    def _undo_poly_vertex(self):
        pts = list(self._draft_pts or [])
        if not pts:
            return
        pts.pop()
        self._draft_pts = pts
        self._hover_snap = None
        self._osnap.clear_tracks()
        self._mark_osnap_dirty()
        self._refresh_summaries()
        if not pts:
            self._set_status(
                u"Polígono: indique el primer vértice (clic en canvas)."
            )
        else:
            self._set_status(
                u"Vértice deshecho · {0} punto(s). Clic para añadir · "
                u"cierre en 1º / Enter (≥3).".format(len(pts))
            )
        self._redraw_plan()

    def _cancel_draft_strip(self):
        had_n = len(self._draft_pts or [])
        self._clear_pick_state(redraw=False)
        self._refresh_summaries()
        if had_n:
            try:
                self._canvas_dbg(u"draft_cancel", draft_n=had_n)
            except Exception:
                pass
            self._set_status(u"Dibujo cancelado.")
        self._redraw_plan()

    def _try_finish_poly(self, status_snap=u""):
        pts = self._normalize_poly_pts_mm(self._draft_pts or [])
        if pts is None:
            self._canvas_dbg(
                u"poly_finish_fail",
                reason=u"normalize",
                draft_n=len(self._draft_pts or []),
            )
            self._set_status(
                u"Polígono incompleto: se necesitan ≥3 vértices distintos."
            )
            return False
        return self._commit_strip_from_pts(pts, snap_lbl=status_snap)

    def _undo_last_strip(self):
        if not self._strips:
            return
        sel = self._selected_strip_idx
        if sel is not None and 0 <= sel < len(self._strips):
            del self._strips[sel]
            self._selected_strip_idx = None
            msg = u"Franja {0} eliminada.".format(sel + 1)
        else:
            self._strips.pop()
            msg = u"Última franja eliminada."
        self._mark_osnap_dirty()
        self._update_armadura_title()
        self._refresh_summaries()
        self._redraw_plan()
        self._set_status(msg)

    def _on_plan_click(self, sender, args):
        # Misma política de pick que Area Rein. Losa Sketch
        if self._panning:
            self._canvas_dbg(u'click_ignored', reason=u'panning')
            try:
                args.Handled = True
            except Exception:
                pass
            return
        try:
            if args.ChangedButton != MouseButton.Left:
                return
        except Exception:
            pass
        cv = self._ui_cv
        if cv is None:
            self._canvas_dbg(u'click_ignored', reason=u'no_canvas')
            return
        pos = None
        try:
            pos = args.GetPosition(cv)
            try:
                args.Handled = True
            except Exception:
                pass
            try:
                cv.Focus()
            except Exception:
                pass
            click_count = 1
            try:
                click_count = int(args.ClickCount)
            except Exception:
                click_count = 1

            pt = self._px_to_mm(pos.X, pos.Y)
            if pt is None:
                self._ensure_scene_transform()
                pt = self._px_to_mm(pos.X, pos.Y)
            if pt is None:
                self._canvas_dbg(
                    u'click_fail',
                    reason=u'px_to_mm',
                    px=u'{0:.1f},{1:.1f}'.format(float(pos.X), float(pos.Y)),
                    **self._canvas_view_snapshot()
                )
                self._set_status(
                    u'No se pudo leer la posición del clic. Espere al dibujo del canvas.'
                )
                return

            mode = getattr(self, u'_draw_mode', DRAW_POLY)
            shift_free = False
            if mode == DRAW_RECT:
                try:
                    shift_free = (
                        (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
                    )
                except Exception:
                    shift_free = False
            sx, sy, kind = self._resolve_snap_mm(
                pt[0], pt[1], raw_only=shift_free
            )
            snap_lbl = osnap_status_label(kind) if kind else u''
            if shift_free:
                snap_lbl = u' · libre'
            pick_pts = list(self._draft_pts or [])
            last_before = pick_pts[-1] if pick_pts else None

            # Borrador vacío: seleccionar franja solo con clic libre (sin OSNAP)
            # dentro del polígono. Si hay snap (end/mid/…), es intención de dibujar
            # — p. ej. esquina de zapata aunque caiga sobre una franja ya cerrada.
            if not pick_pts:
                hit = None
                if kind is None:
                    hit = self._hit_strip_at_mm(pt)
                if hit is not None:
                    self._select_strip(hit)
                    self._canvas_click_diag(
                        pos,
                        pt,
                        sx,
                        sy,
                        kind,
                        u'select_strip',
                        click_count=click_count,
                        strip=hit + 1,
                    )
                    return
                if self._selected_strip_idx is not None and kind is None:
                    # Clic libre fuera: deseleccionar y seguir dibujando
                    self._clear_strip_selection(
                        status=u"Sin franja seleccionada · dibuje o elija otra."
                    )

            if mode == DRAW_POLY:
                if len(pick_pts) >= 3 and self._point_near_first_vertex((sx, sy)):
                    ok = self._try_finish_poly(status_snap=snap_lbl)
                    self._canvas_click_diag(
                        pos,
                        pt,
                        sx,
                        sy,
                        kind,
                        u'close_ok' if ok else u'close_fail',
                        click_count=click_count,
                    )
                    return

                if click_count >= 2 and len(pick_pts) >= 2:
                    appended = False
                    if pick_pts:
                        dx = float(sx) - float(pick_pts[-1][0])
                        dy = float(sy) - float(pick_pts[-1][1])
                        if dx * dx + dy * dy >= (
                            float(_POLY_MIN_EDGE_MM) * float(_POLY_MIN_EDGE_MM)
                        ):
                            pick_pts.append((float(sx), float(sy)))
                            self._draft_pts = pick_pts
                            self._hover_snap = None
                            self._mark_osnap_dirty()
                            appended = True
                    ok = self._try_finish_poly(status_snap=snap_lbl)
                    self._canvas_click_diag(
                        pos,
                        pt,
                        sx,
                        sy,
                        kind,
                        u'double_close_ok' if ok else u'double_close_fail',
                        click_count=click_count,
                        dbl_append=1 if appended else 0,
                    )
                    return

                if pick_pts:
                    dx = float(sx) - float(pick_pts[-1][0])
                    dy = float(sy) - float(pick_pts[-1][1])
                    dist_last = dx * dx + dy * dy
                    if dist_last < (
                        float(_POLY_MIN_EDGE_MM) * float(_POLY_MIN_EDGE_MM)
                    ):
                        self._canvas_click_diag(
                            pos,
                            pt,
                            sx,
                            sy,
                            kind,
                            u'reject_too_close',
                            click_count=click_count,
                            min_edge_mm=_POLY_MIN_EDGE_MM,
                            dist_last_mm=u'{0:.2f}'.format(dist_last ** 0.5),
                        )
                        self._set_status(
                            u'Vértice demasiado cercano al anterior. Elija otro punto.'
                        )
                        return

                pick_pts.append((float(sx), float(sy)))
                self._draft_pts = pick_pts
                self._hover_snap = None
                self._mark_osnap_dirty()
                if self._selected_strip_idx is not None and len(pick_pts) == 1:
                    self._selected_strip_idx = None
                    self._update_armadura_title()
                self._refresh_summaries()
                self._redraw_plan()
                n = len(pick_pts)
                self._canvas_click_diag(
                    pos,
                    pt,
                    sx,
                    sy,
                    kind,
                    u'vertex_added',
                    last_pt=last_before,
                    click_count=click_count,
                    vertex_n=n,
                    shift_free=0,
                )
                if n == 1:
                    self._set_status(
                        u'Vértice 1 ({0:.0f}, {1:.0f}) mm{2}. Siguiente vértice…'.format(
                            sx, sy, snap_lbl
                        )
                    )
                elif n < 3:
                    self._set_status(
                        u'Vértice {0} ({1:.0f}, {2:.0f}){3}. Faltan ≥{4} para cerrar.'.format(
                            n, sx, sy, snap_lbl, 3 - n
                        )
                    )
                else:
                    self._set_status(
                        u'Vértice {0}{1}. Clic en el 1º (verde), Enter o doble clic '
                        u'para cerrar · Retroceso deshace.'.format(n, snap_lbl)
                    )
                return

            # --- Rectángulo 2 puntos ---
            if not pick_pts:
                self._draft_pts = [(float(sx), float(sy))]
                self._hover_snap = None
                self._mark_osnap_dirty()
                if self._selected_strip_idx is not None:
                    self._selected_strip_idx = None
                    self._update_armadura_title()
                self._refresh_summaries()
                self._redraw_plan()
                self._canvas_click_diag(
                    pos, pt, sx, sy, kind, u'rect_corner_a', click_count=click_count
                )
                self._set_status(
                    u'Esquina A ({0:.0f}, {1:.0f}) mm{2}. Indique la esquina opuesta.'.format(
                        sx, sy, snap_lbl
                    )
                )
                return

            pts = None
            try:
                pts = rect_from_two_points_mm(pick_pts[0], (sx, sy))
            except Exception:
                pts = None
            if not pts:
                self._canvas_click_diag(
                    pos, pt, sx, sy, kind, u'rect_fail_short', click_count=click_count
                )
                self._set_status(
                    u'Distancia insuficiente. Mantenga esquina A e indique otra B.'
                )
                return
            ok = self._commit_strip_from_pts(pts, snap_lbl=snap_lbl)
            self._canvas_click_diag(
                pos,
                pt,
                sx,
                sy,
                kind,
                u'rect_commit' if ok else u'rect_commit_fail',
                click_count=click_count,
            )
        except Exception as ex:
            try:
                fields = self._canvas_view_snapshot()
                if pos is not None:
                    fields[u'px'] = u'{0:.1f},{1:.1f}'.format(
                        float(pos.X), float(pos.Y)
                    )
                fields[u'error'] = _as_unicode(ex)
                self._canvas_dbg(u'click_exception', **fields)
            except Exception:
                pass
            try:
                self._set_status(u'Error al añadir vértice: {0}'.format(_as_unicode(ex)))
            except Exception:
                pass

    def _update_plan_cursor(self):
        cv = self._ui_cv
        if cv is None:
            return
        try:
            if self._panning:
                cv.Cursor = Cursors.SizeAll
            else:
                cv.Cursor = Cursors.Cross
        except Exception:
            pass

    def _end_plan_pan(self, restore_cursor=True):
        was = bool(self._panning)
        self._panning = False
        cv = self._ui_cv
        if cv is not None:
            try:
                if cv.IsMouseCaptured:
                    cv.ReleaseMouseCapture()
            except Exception:
                pass
            if restore_cursor:
                self._update_plan_cursor()
        return was

    def _reset_plan_view(self):
        self._end_plan_pan(restore_cursor=True)
        self._view_zoom = 1.0
        self._view_pan_x = 0.0
        self._view_pan_y = 0.0
        self._redraw_plan()
        self._set_status(u"Vista restablecida (fit).")

    def _on_plan_wheel(self, sender, args):
        """Zoom hacia el cursor (misma lógica que Area Rein. Losa Sketch)."""
        cv = self._ui_cv
        sb = self._ensure_scene_transform()
        if cv is None or sb is None:
            return
        try:
            delta = int(args.Delta)
        except Exception:
            return
        if delta == 0:
            return

        zoom_min = 0.25
        zoom_max = 16.0
        step = 1.06
        zoom = float(self._view_zoom) if self._view_zoom else 1.0
        zoom = max(zoom_min, min(zoom_max, zoom))
        try:
            zoom_new = zoom * math.pow(step, float(delta) / 120.0)
        except Exception:
            zoom_new = zoom * (step if delta > 0 else (1.0 / step))
        zoom_new = max(zoom_min, min(zoom_max, zoom_new))
        if abs(zoom_new - zoom) < 1e-12:
            try:
                args.Handled = True
            except Exception:
                pass
            return

        try:
            pos = args.GetPosition(cv)
            mx = float(pos.X)
            my = float(pos.Y)
            cw, ch = self._canvas_size(cv)
        except Exception:
            return
        if cw < 40 or ch < 40:
            return

        try:
            scale = float(sb[u"scale"])
            if scale < 1e-12:
                return
            min_x = float(sb[u"min_x"])
            max_y = float(sb[u"max_y"])
            max_x = float(sb.get(u"max_x", min_x))
            min_y = float(sb.get(u"min_y", max_y))
            ox = float(sb[u"ox"])
            oy = float(sb[u"oy"])
        except Exception:
            return

        xmm = min_x + (mx - ox) / scale
        ymm = max_y - (my - oy) / scale
        actual_factor = zoom_new / zoom
        scale_new = scale * actual_factor
        cx_mm = xmm + (cw / 2.0 - mx) / scale_new
        cy_mm = ymm - (ch / 2.0 - my) / scale_new
        bbox_cx = (min_x + max_x) / 2.0
        bbox_cy = (min_y + max_y) / 2.0
        self._view_pan_x = cx_mm - bbox_cx
        self._view_pan_y = cy_mm - bbox_cy
        self._view_zoom = zoom_new
        try:
            args.Handled = True
        except Exception:
            pass
        self._redraw_plan()

    def _on_plan_middle_down(self, sender, args):
        try:
            if args.ChangedButton != MouseButton.Middle:
                return
        except Exception:
            return
        cv = self._ui_cv
        if cv is None:
            return
        if self._ensure_scene_transform() is None:
            try:
                args.Handled = True
            except Exception:
                pass
            return
        try:
            pos = args.GetPosition(cv)
            self._pan_last_x = float(pos.X)
            self._pan_last_y = float(pos.Y)
        except Exception:
            return
        self._panning = True
        try:
            cv.CaptureMouse()
        except Exception:
            pass
        self._update_plan_cursor()
        try:
            args.Handled = True
        except Exception:
            pass

    def _on_plan_middle_up(self, sender, args):
        try:
            if args.ChangedButton != MouseButton.Middle:
                return
        except Exception:
            return
        self._end_plan_pan(restore_cursor=True)
        try:
            args.Handled = True
        except Exception:
            pass

    def _on_plan_lost_capture(self, sender, args):
        self._end_plan_pan(restore_cursor=True)

    def _on_plan_enter(self, sender, args):
        self._mouse_in_canvas = True

    def _on_plan_leave(self, sender, args):
        self._mouse_in_canvas = False
        if self._hover_snap is not None:
            self._hover_snap = None
            self._redraw_plan()

    def _on_plan_move(self, sender, args):
        if self._panning and self._scene_base is not None:
            try:
                pos = args.GetPosition(self._ui_cv)
                mx = float(pos.X)
                my = float(pos.Y)
                dx_px = mx - float(self._pan_last_x)
                dy_px = my - float(self._pan_last_y)
                self._pan_last_x = mx
                self._pan_last_y = my
                scale = float(self._scene_base.get(u"scale") or 0.0)
                if scale > 1e-12 and (abs(dx_px) > 1e-9 or abs(dy_px) > 1e-9):
                    self._view_pan_x = float(self._view_pan_x or 0.0) - (
                        dx_px / scale
                    )
                    self._view_pan_y = float(self._view_pan_y or 0.0) + (
                        dy_px / scale
                    )
                    self._redraw_plan()
            except Exception:
                pass
            try:
                args.Handled = True
            except Exception:
                pass
            return
        if not self._canvas_mouse_active():
            return
        # Hover snap preview (misma lógica que Area Rein. Losa Sketch)
        try:
            pos = args.GetPosition(self._ui_cv)
            pt = self._px_to_mm(pos.X, pos.Y)
            if pt is None:
                return
            n_acq_before = len(getattr(self._osnap, u"_acquired_tracks", None) or [])
            sx, sy, kind = self._resolve_snap_mm(pt[0], pt[1])
            n_acq = len(getattr(self._osnap, u"_acquired_tracks", None) or [])
            n_ot = len(getattr(self._osnap, u"_ot_points", None) or [])
            has_acq = n_acq > 0 or n_ot > 0
            pick_pts = self._draft_pts or []
            if kind is None and not pick_pts and not has_acq:
                if self._hover_snap is not None:
                    self._hover_snap = None
                    self._redraw_plan()
                return
            new_hover = (sx, sy, kind)
            prev = self._hover_snap
            force = n_acq != n_acq_before or n_ot != getattr(self, u"_prev_ot_n", -1)
            self._prev_ot_n = n_ot
            if (
                not force
                and prev is not None
                and abs(float(prev[0]) - float(new_hover[0])) < 0.05
                and abs(float(prev[1]) - float(new_hover[1])) < 0.05
                and prev[2] == new_hover[2]
            ):
                if not pick_pts:
                    return
            self._hover_snap = new_hover
            changed = bool(pick_pts) or prev != self._hover_snap
            if not changed and prev is not None and self._hover_snap is not None:
                changed = (
                    abs(float(prev[0]) - float(self._hover_snap[0])) > 0.05
                    or abs(float(prev[1]) - float(self._hover_snap[1])) > 0.05
                    or prev[2] != self._hover_snap[2]
                )
            if changed:
                self._redraw_plan()
            if self._canvas_inst.is_ui_enabled():
                now = time.time()
                hover_key = (
                    round(float(sx), 1),
                    round(float(sy), 1),
                    kind,
                )
                if hover_key != self._hover_log_key or (now - self._hover_log_at) > 0.45:
                    self._hover_log_key = hover_key
                    self._hover_log_at = now
                    fields = self._canvas_view_snapshot()
                    fields[u"px"] = u"{0:.1f},{1:.1f}".format(
                        float(pos.X), float(pos.Y)
                    )
                    fields[u"raw_mm"] = u"{0:.1f},{1:.1f}".format(
                        float(pt[0]), float(pt[1])
                    )
                    fields[u"snap_mm"] = u"{0:.1f},{1:.1f}".format(float(sx), float(sy))
                    fields[u"snap_kind"] = kind or u"none"
                    dmm = dist_mm(pt, (sx, sy))
                    if dmm is not None:
                        fields[u"snap_delta_mm"] = u"{0:.2f}".format(float(dmm))
                    self._canvas_dbg(u"hover", **fields)
        except Exception:
            pass

    def _on_key_down(self, sender, args):
        try:
            key = args.Key
        except Exception:
            return
        # No interceptar teclas mientras se edita un TextBox del rail (sep, etc.)
        if self._keyboard_focus_in_textbox():
            return
        try:
            mods = Keyboard.Modifiers
            ctrl_shift = (
                (mods & ModifierKeys.Control) == ModifierKeys.Control
                and (mods & ModifierKeys.Shift) == ModifierKeys.Shift
            )
            if ctrl_shift and key == Key.D:
                self._toggle_canvas_debug()
                args.Handled = True
                return
            if ctrl_shift and key == Key.L:
                self._open_canvas_debug_log()
                args.Handled = True
                return
            if key == Key.D0 or key == Key.NumPad0:
                if (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control:
                    self._reset_plan_view()
                    args.Handled = True
                    return
            if key == Key.Enter or key == Key.Return:
                if (
                    getattr(self, u"_draw_mode", DRAW_POLY) == DRAW_POLY
                    and len(self._draft_pts or []) >= 3
                ):
                    try:
                        self._canvas_dbg(
                            u"enter_close",
                            draft_n=len(self._draft_pts or []),
                        )
                    except Exception:
                        pass
                    if self._try_finish_poly(status_snap=u""):
                        args.Handled = True
                return
            if key == Key.Back:
                if self._draft_pts:
                    self._undo_poly_vertex()
                    args.Handled = True
                return
            if key == Key.Delete:
                if self._draft_pts:
                    self._undo_poly_vertex()
                    args.Handled = True
                elif self._strips:
                    self._undo_last_strip()
                    args.Handled = True
                return
            if key == Key.Escape:
                if self._draft_pts:
                    self._cancel_draft_strip()
                    args.Handled = True
        except Exception:
            pass

    def _open_manual(self, sender=None, args=None):
        path = None
        try:
            import bimtools_paths

            pb = bimtools_paths.get_pushbutton_dir()
            if pb:
                cand = os.path.join(pb, u"manual_usuario.html")
                if os.path.isfile(cand):
                    path = cand
        except Exception:
            pass
        if path is None:
            try:
                ext = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                for tab in os.listdir(ext):
                    if not tab.endswith(u".tab"):
                        continue
                    panel = os.path.join(ext, tab, u"3D Rebar.panel")
                    if not os.path.isdir(panel):
                        continue
                    for pb in os.listdir(panel):
                        if u"FundacionCorridaFranja" not in pb:
                            continue
                        cand = os.path.join(panel, pb, u"manual_usuario.html")
                        if os.path.isfile(cand):
                            path = cand
                            break
            except Exception:
                pass
        if not path:
            _mostrar_aviso(self._uiapp, u"No se encontró manual_usuario.html.")
            return
        try:
            os.startfile(path)
        except Exception as ex:
            _mostrar_aviso(self._uiapp, u"No se pudo abrir el manual.", _as_unicode(ex))

    def _dispose_col_event(self):
        try:
            if self._col_event is not None:
                self._col_event.Dispose()
        except Exception:
            pass
        self._col_event = None

    def _on_colocar_click(self, sender=None, args=None):
        if not self._strips:
            _mostrar_aviso(
                self._uiapp,
                u"Dibuje al menos una franja antes de colocar.",
            )
            return
        # Persistir rail → franja seleccionada antes de cerrar la UI
        self._sync_rail_to_selected_strip()
        self._colocar_pending = True
        self._col_handler._ctrl = self
        ev = self._col_event
        if ev is None:
            self._colocar_pending = False
            self._col_handler._ctrl = None
            _mostrar_aviso(
                self._uiapp,
                u"No se pudo iniciar la colocación.",
                content=u"ExternalEvent no disponible.",
            )
            return
        try:
            ev.Raise()
        except Exception as ex:
            self._colocar_pending = False
            self._col_handler._ctrl = None
            _mostrar_aviso(
                self._uiapp,
                u"No se pudo iniciar la colocación.",
                content=_as_unicode(ex),
            )
            return
        # Cerrar UI de inmediato; Execute modela con los datos ya en memoria
        self._close()

    def _execute_colocar(self, uiapp):
        doc = self._doc
        uidoc = self._uidoc
        view = self._source_view
        if doc is None:
            return
        entries = self._entries
        hosts_prev = list(zip(self._hosts, self._previews))
        avisos = []
        rebars_all = []
        n_ok = 0
        n_strips = len(self._strips or [])
        active_view_for_params = None
        try:
            if uidoc is not None:
                active_view_for_params = uidoc.ActiveView
        except Exception:
            active_view_for_params = None
        if active_view_for_params is None:
            active_view_for_params = view
        sheet_number = _sheet_number_from_view(active_view_for_params)
        if not sheet_number:
            sheet_number = _sheet_number_from_view(view)
        conjunto_guid = iniciar_armadura_conjunto_guid_ejecucion()
        try:
            with _FundacionFranjaColocarProgress(n_strips, _DIALOG_TITLE) as pbar:
                t = Transaction(doc, u"Arainco: Fundación corrida franja")
                t.Start()
                try:
                    for idx, strip in enumerate(self._strips):
                        try:
                            pbar.update(
                                idx + 1,
                                label=u"Franja {0}".format(idx + 1),
                            )
                        except Exception:
                            pass
                        cfg = strip.get(u"config") or self._default_rail_config()
                        bt_tr, e1 = self._resolver_bar_type_from_idx(
                            doc, entries, cfg.get(u"trans_idx", 0)
                        )
                        bt_lo, e2 = self._resolver_bar_type_from_idx(
                            doc, entries, cfg.get(u"long_idx", 0)
                        )
                        if bt_tr is None:
                            avisos.append(
                                u"Franja {0}: {1}".format(
                                    idx + 1, e1 or u"Tipo transversal no válido."
                                )
                            )
                            continue
                        if bt_lo is None:
                            avisos.append(
                                u"Franja {0}: {1}".format(
                                    idx + 1, e2 or u"Tipo longitudinal no válido."
                                )
                            )
                            continue
                        try:
                            trans_sep = int(
                                cfg.get(u"trans_sep_mm") or _FRANJA_SEP_MM_DEFAULT
                            )
                        except Exception:
                            trans_sep = int(_FRANJA_SEP_MM_DEFAULT)
                        try:
                            long_sep = int(
                                cfg.get(u"long_sep_mm") or _FRANJA_SEP_MM_DEFAULT
                            )
                        except Exception:
                            long_sep = int(_FRANJA_SEP_MM_DEFAULT)
                        trans_sep = _clamp_franja_sep_mm(trans_sep)
                        long_sep = _clamp_franja_sep_mm(long_sep)
                        grade = cfg.get(u"grade") or _ewf._DOSIFICACION_HORMIGON_DEFAULT
                        d_long = _rebar_nominal_diameter_mm(bt_lo) or 0.0
                        d_tr = _rebar_nominal_diameter_mm(bt_tr) or 0.0
                        wf = assign_host_to_strip(hosts_prev, strip)
                        if wf is None or not isinstance(wf, WallFoundation):
                            avisos.append(
                                u"Franja {0}: sin Wall Foundation asociada.".format(
                                    idx + 1
                                )
                            )
                            continue
                        joined = _ewf._wf_collect_joined_element_ids(doc, wf)
                        _ewf._wf_unjoin_all(doc, wf, joined, avisos)
                        try:
                            doc.Regenerate()
                        except Exception:
                            pass
                        wf = doc.GetElement(wf.Id)
                        if wf is None:
                            avisos.append(
                                u"Franja {0}: host inválido tras desunir.".format(
                                    idx + 1
                                )
                            )
                            continue
                        base_geo, hint = _ewf._geometria_wall_foundation_inferior(
                            wf, d_long, d_tr
                        )
                        if hint:
                            avisos.append(u"Franja {0}: {1}".format(idx + 1, hint))
                        if base_geo is None:
                            # Fallback: solo franja a Z de bbox
                            z0, z1 = _ewf._wf_z_range_ft(wf)
                            z_ref = 0.5 * ((z0 or 0.0) + (z1 or 0.0))
                            from fundacion_corrida_franja_geom import (
                                build_lines_from_strip_ft,
                            )

                            ll, wl, uw = build_lines_from_strip_ft(
                                strip, z_ref, view_frame=self._view_frame
                            )
                            if ll is None or wl is None:
                                avisos.append(
                                    u"Franja {0}: no se resolvió geometría.".format(
                                        idx + 1
                                    )
                                )
                                _ewf._wf_rejoin_all(doc, wf, joined, avisos)
                                continue
                            base_geo = {
                                "long_line": ll,
                                "width_line": wl,
                                "usable_w_ft": uw,
                                "n_cara": XYZ.BasisZ.Negate(),
                                "z0": z0,
                                "z1": z1,
                            }
                        geo = merge_strip_into_host_geo(
                            base_geo, strip, view_frame=self._view_frame
                        )
                        if geo is None:
                            avisos.append(
                                u"Franja {0}: no se pudo fusionar franja con host.".format(
                                    idx + 1
                                )
                            )
                            _ewf._wf_rejoin_all(doc, wf, joined, avisos)
                            continue
                        geo = _ewf._wf_apply_recubrimiento_ejes_franja(
                            wf, geo, d_long, d_tr
                        )
                        if geo is None or geo.get("long_line") is None:
                            avisos.append(
                                u"Franja {0}: recubrimiento dejó geometría inválida.".format(
                                    idx + 1
                                )
                            )
                            _ewf._wf_rejoin_all(doc, wf, joined, avisos)
                            continue
                        L_mm = float(strip.get(u"length_mm") or 0)
                        needs_lap = L_mm > float(_ewf._MAX_STOCK_MM) + 0.01
                        if needs_lap:
                            try:
                                max_mm = float(
                                    cfg.get(u"max_bar_mm") or _ewf._MAX_STOCK_MM
                                )
                            except Exception:
                                max_mm = float(_ewf._MAX_STOCK_MM)
                            try:
                                lap_mm = float(
                                    cfg.get(u"lap_mm") or _ewf._LAP_DEFAULT_MM
                                )
                            except Exception:
                                lap_mm = float(_ewf._LAP_DEFAULT_MM)
                            if lap_mm <= 0:
                                lap_mm = _ewf._wf_traslape_mm_longitudinal(
                                    d_long, None, grade
                                )
                            if max_mm <= lap_mm + 1.0:
                                avisos.append(
                                    u"Franja {0}: largo máx. debe ser > empalme.".format(
                                        idx + 1
                                    )
                                )
                                _ewf._wf_rejoin_all(doc, wf, joined, avisos)
                                continue
                        else:
                            max_mm = float(_ewf._MAX_STOCK_MM)
                            lap_mm = float(_ewf._LAP_DEFAULT_MM)
                        created_u = []
                        n_t = _ewf._colocar_trans_u(
                            doc,
                            wf,
                            bt_tr,
                            trans_sep,
                            geo,
                            avisos,
                            rebars_out=created_u,
                            concrete_grade=grade,
                        )
                        long_axis = geo.get("long_line")
                        if n_t > 0 and d_tr > 1e-6:
                            try:
                                long_axis = (
                                    _ewf._wf_traslada_linea_hacia_interior_hormigon_mm(
                                        long_axis,
                                        geo.get("n_cara"),
                                        float(d_tr),
                                    )
                                )
                            except Exception:
                                long_axis = geo.get("long_line")
                            if long_axis is None:
                                long_axis = geo.get("long_line")
                        created_l = []
                        n_l = _ewf._colocar_rebar_en_host(
                            doc,
                            wf,
                            bt_lo,
                            long_sep,
                            long_axis,
                            geo.get("width_line"),
                            geo.get("usable_w_ft"),
                            needs_lap,
                            max_mm,
                            lap_mm,
                            avisos,
                            geo=geo,
                            rebars_out=created_l,
                            concrete_grade=grade,
                            active_view=view,
                        )
                        nivel_nombre = _nivel_nombre_wall_foundation(doc, wf)
                        # Cara inferior (F): U = luz menor (i); long = luz mayor (s)
                        _stamp_rebars_franja(
                            created_u,
                            conjunto_guid,
                            nivel_nombre,
                            sheet_number,
                            ARMADURA_UBICACION_INFERIOR,
                            _ARMADURA_POSICION_TRANS,
                        )
                        _stamp_rebars_franja(
                            created_l,
                            conjunto_guid,
                            nivel_nombre,
                            sheet_number,
                            ARMADURA_UBICACION_INFERIOR,
                            _ARMADURA_POSICION_LONG,
                        )
                        _ewf._wf_rejoin_all(doc, wf, joined, avisos)
                        if n_t > 0 or n_l > 0:
                            n_ok += 1
                            rebars_all.extend(created_u)
                            rebars_all.extend(created_l)
                            avisos.append(
                                u"Franja {0}: U×{1} + long×{2}.".format(
                                    idx + 1, n_t, n_l
                                )
                            )
                    try:
                        doc.Regenerate()
                    except Exception:
                        pass
                    if rebars_all:
                        try:
                            pbar.update(
                                n_strips,
                                label=u"Etiquetas / MRA",
                            )
                        except Exception:
                            pass
                        # Etiquetas + MRA antes de «solo barra central» (la API de tag
                        # suele requerir el conjunto completo visible).
                        tag_view = active_view_for_params
                        if tag_view is None:
                            tag_view = view
                        if tag_view is not None:
                            try:
                                doc.Regenerate()
                            except Exception:
                                pass
                            try:
                                n_tags = _ewf._wf_etiquetar_rebar_sets_independent_tag(
                                    doc,
                                    tag_view,
                                    rebars_all,
                                    avisos,
                                )
                                if n_tags > 0:
                                    avisos.append(
                                        u"Etiquetas «{0}» (tipo = RebarShape): {1} creada(s).".format(
                                            _ewf._WF_REBAR_TAG_FAMILY_NAME,
                                            int(n_tags),
                                        )
                                    )
                            except Exception as ex_tag:
                                avisos.append(
                                    u"Etiquetas no aplicadas: {0}".format(
                                        _as_unicode(ex_tag)
                                    )
                                )
                            try:
                                from geometria_estribos_viga import (
                                    crear_multi_rebar_annotations_por_nombre_tipo,
                                )

                                n_mra = crear_multi_rebar_annotations_por_nombre_tipo(
                                    doc,
                                    tag_view,
                                    rebars_all,
                                    avisos,
                                    _ewf._WF_MULTI_REBAR_ANNOTATION_TYPE_NAME,
                                )
                                if n_mra > 0:
                                    avisos.append(
                                        u"Multi-Rebar Annotation «{0}»: {1} creada(s).".format(
                                            _ewf._WF_MULTI_REBAR_ANNOTATION_TYPE_NAME,
                                            int(n_mra),
                                        )
                                    )
                            except Exception as ex_mra:
                                avisos.append(
                                    u"Multi-Rebar Annotation no aplicada: {0}".format(
                                        _as_unicode(ex_mra)
                                    )
                                )
                        try:
                            pbar.update(
                                n_strips,
                                label=u"Visibilidad",
                            )
                        except Exception:
                            pass
                        views_unob = []
                        active_view = active_view_for_params
                        if active_view is None:
                            active_view = view
                        for v_cand in (active_view, view):
                            if v_cand is None:
                                continue
                            try:
                                vid = element_id_to_int(v_cand.Id)
                            except Exception:
                                vid = id(v_cand)
                            if any(x[0] == vid for x in views_unob):
                                continue
                            views_unob.append((vid, v_cand))
                        for _vid, v_unob in views_unob:
                            if _ewf._wf_vista_es_planta(v_unob):
                                try:
                                    _ewf._wf_aplicar_presentacion_solo_barra_central_planta(
                                        v_unob, rebars_all
                                    )
                                except Exception:
                                    pass
                            try:
                                n_unob = _ewf._wf_aplicar_unobscured_rebars(
                                    doc, v_unob, rebars_all
                                )
                                if n_unob > 0:
                                    avisos.append(
                                        u"View Unobscured (+ sólido): {0} barra(s) en «{1}».".format(
                                            int(n_unob),
                                            _as_unicode(
                                                getattr(v_unob, u"Name", None)
                                                or u"vista"
                                            ),
                                        )
                                    )
                            except Exception as ex_unob:
                                avisos.append(
                                    u"View Unobscured no aplicado: {0}".format(
                                        _as_unicode(ex_unob)
                                    )
                                )
                    t.Commit()
                except Exception as ex:
                    try:
                        t.RollBack()
                    except Exception:
                        pass
                    _mostrar_aviso(
                        uiapp,
                        u"Error (se revirtió la transacción).",
                        content=_as_unicode(ex),
                    )
                    return
        finally:
            try:
                finalizar_armadura_conjunto_guid_ejecucion()
            except Exception:
                pass
        msg = u"Franjas con armadura: {0} / {1}.".format(n_ok, len(self._strips))
        try:
            self._set_status(msg)
        except Exception:
            pass

    def _wire(self):
        from System.Windows import SizeChangedEventHandler
        from System.Windows.Controls import SelectionChangedEventHandler

        win = self._win
        self._ui_cv = win.FindName(u"CvPlan")
        self._ui_btn_draw_rect = win.FindName(u"BtnDrawRect")
        self._ui_btn_draw_poly = win.FindName(u"BtnDrawPoly")
        self._ui_txt_canvas_debug = win.FindName(u"TxtCanvasDebug")
        if self._ui_btn_draw_rect is not None:
            self._ui_btn_draw_rect.Click += RoutedEventHandler(
                lambda s, e: self._set_draw_mode(DRAW_RECT)
            )
        if self._ui_btn_draw_poly is not None:
            self._ui_btn_draw_poly.Click += RoutedEventHandler(
                lambda s, e: self._set_draw_mode(DRAW_POLY)
            )
        win.FindName(u"BtnColocar").Click += RoutedEventHandler(self._on_colocar_click)
        win.FindName(u"BtnCancelar").Click += RoutedEventHandler(
            lambda s, e: self._close()
        )
        win.FindName(u"BtnManual").Click += RoutedEventHandler(self._open_manual)
        cv = self._ui_cv
        if cv is not None:
            try:
                cv.Focusable = True
                # Sin Background el Canvas no recibe hit-test en zonas vacías (Losa Sketch)
                cv.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
                self._update_plan_cursor()
            except Exception:
                pass
            cv.MouseEnter += MouseEventHandler(self._on_plan_enter)
            cv.MouseLeave += MouseEventHandler(self._on_plan_leave)
            try:
                cv.PreviewMouseLeftButtonDown += MouseButtonEventHandler(
                    self._on_plan_click
                )
            except Exception:
                cv.MouseLeftButtonDown += MouseButtonEventHandler(self._on_plan_click)
            try:
                cv.PreviewMouseDown += MouseButtonEventHandler(
                    self._on_plan_middle_down
                )
            except Exception:
                cv.MouseDown += MouseButtonEventHandler(self._on_plan_middle_down)
            try:
                cv.PreviewMouseUp += MouseButtonEventHandler(self._on_plan_middle_up)
            except Exception:
                cv.MouseUp += MouseButtonEventHandler(self._on_plan_middle_up)
            try:
                cv.LostMouseCapture += MouseEventHandler(self._on_plan_lost_capture)
            except Exception:
                pass
            cv.MouseMove += MouseEventHandler(self._on_plan_move)
            try:
                cv.PreviewMouseWheel += MouseWheelEventHandler(self._on_plan_wheel)
            except Exception:
                cv.MouseWheel += MouseWheelEventHandler(self._on_plan_wheel)
        try:
            win.Focusable = True
            win.PreviewKeyDown += KeyEventHandler(self._on_key_down)
        except Exception:
            win.KeyDown += KeyEventHandler(self._on_key_down)
        win.SizeChanged += SizeChangedEventHandler(
            lambda s, e: self._redraw_plan()
        )
        win.Closed += EventHandler(self._on_closed)
        for name in (u"TxtTransSep", u"TxtLongSep"):
            tb = win.FindName(name)
            if tb is not None:
                tb.LostFocus += RoutedEventHandler(
                    lambda s, a, tbx=tb: self._on_sep_lost_focus(s, a, tbx)
                )
        for tb_name, up_name, dn_name in (
            (u"TxtTransSep", u"BtnTransSepUp", u"BtnTransSepDown"),
            (u"TxtLongSep", u"BtnLongSepUp", u"BtnLongSepDown"),
        ):
            tb = win.FindName(tb_name)
            bup = win.FindName(up_name)
            bdn = win.FindName(dn_name)
            if bup is not None and tb is not None:
                bup.Click += RoutedEventHandler(
                    lambda s, e, tbx=tb: self._step_franja_sep(
                        tbx, _FRANJA_SEP_MM_STEP
                    )
                )
            if bdn is not None and tb is not None:
                bdn.Click += RoutedEventHandler(
                    lambda s, e, tbx=tb: self._step_franja_sep(
                        tbx, -_FRANJA_SEP_MM_STEP
                    )
                )
        tmax = win.FindName(u"TxtMaxBarMm")
        tlap = win.FindName(u"TxtLapMm")
        if tmax is not None:
            tmax.LostFocus += RoutedEventHandler(self._on_max_lost_focus)
        if tlap is not None:
            tlap.LostFocus += RoutedEventHandler(self._on_lap_lost_focus)
        cmb_long = win.FindName(u"CmbLongDiam")
        if cmb_long is not None:
            cmb_long.SelectionChanged += SelectionChangedEventHandler(
                self._on_long_diam_changed
            )
        cmb_tr = win.FindName(u"CmbTransDiam")
        if cmb_tr is not None:
            cmb_tr.SelectionChanged += SelectionChangedEventHandler(
                self._on_rail_changed
            )
        cmb_dos = win.FindName(u"CmbDosificacionHormigon")
        if cmb_dos is not None:
            cmb_dos.SelectionChanged += SelectionChangedEventHandler(
                self._on_rail_changed
            )

    def _sync_lap(self):
        if self._rail_loading:
            return
        try:
            cmb = self._win.FindName(u"CmbLongDiam")
            tlap = self._win.FindName(u"TxtLapMm")
            bt, _ = _ewf._resolver_bar_type_from_combo(self._doc, cmb, self._entries)
            d = _rebar_nominal_diameter_mm(bt) if bt else None
            if d is None or tlap is None:
                return
            gr = _ewf._read_dosificacion_hormigon(
                self._win.FindName(u"CmbDosificacionHormigon")
            )
            v = traslape_mm_from_nominal_diameter_mm(float(d), gr)
            if v is not None:
                tlap.Text = _as_unicode(int(round(v)))
        except Exception:
            pass

    def _on_closed(self, sender, args):
        _clear_singleton()
        try:
            self._win = None
        except Exception:
            pass
        # Colocación pendiente: no Dispose del ExternalEvent hasta Execute
        if getattr(self, u"_colocar_pending", False):
            return
        self._dispose_col_event()

    def _close(self):
        try:
            if self._win is not None:
                self._win.Close()
        except Exception:
            pass

    def show(self):
        self._win = XamlReader.Parse(_XAML)
        self._wire()
        try:
            view_name = getattr(self._source_view, u"Name", u"") or u""
        except Exception:
            view_name = u""
        self._canvas_inst.begin_session(
            {
                u"view": _as_unicode(view_name),
                u"hosts": len(self._hosts or []),
                u"log": self._canvas_inst.log_path,
            }
        )
        self._refresh_canvas_debug_ui()
        self._cargar_combos()
        self._mark_osnap_dirty()
        self._apply_draw_mode_visuals()
        self._refresh_summaries()
        hwnd = None
        try:
            hwnd = revit_main_hwnd(self._uiapp)
        except Exception:
            pass
        try:
            from System.Windows.Interop import WindowInteropHelper

            if hwnd:
                WindowInteropHelper(self._win).Owner = hwnd
        except Exception:
            pass
        position_wpf_window_top_left_at_active_view(self._win, self._uidoc, hwnd)
        try:
            self._win.WindowState = WindowState.Maximized
        except Exception:
            pass
        _set_singleton(self._win)
        self._win.Show()
        try:
            self._win.UpdateLayout()
        except Exception:
            pass
        try:
            if self._ui_cv is not None and bool(self._ui_cv.IsMouseOver):
                self._mouse_in_canvas = True
        except Exception:
            pass
        self._redraw_plan()
        self._draft_pts = []
        self._hover_snap = None
        self._osnap.clear_tracks()
        self._mark_osnap_dirty()
        self._set_status(
            u"Modo Polígono (por defecto): clic = vértice · OSNAP · "
            u"Rueda = zoom · clic rueda = pan · Ctrl+0 = reset vista."
        )
        try:
            if self._ui_cv is not None:
                self._ui_cv.Focus()
        except Exception:
            pass


def run(revit):
    uiapp = revit
    try:
        uidoc = uiapp.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return
    doc = uidoc.Document
    if doc is None or doc.IsFamilyDocument:
        _mostrar_aviso(uiapp, u"Abra un proyecto (no un family document).")
        return
    try:
        active_view = uidoc.ActiveView
    except Exception:
        active_view = None
    if not _vista_es_planta(active_view):
        _mostrar_aviso(
            uiapp,
            u"Esta herramienta solo funciona en vistas de planta.",
            content=u"Abra una planta y vuelva a ejecutar.",
        )
        return

    existing = _get_singleton()
    if existing is not None:
        try:
            existing.WindowState = WindowState.Maximized
            existing.Show()
            existing.Activate()
            _mostrar_aviso(uiapp, u"La herramienta ya está en ejecución.")
            return
        except Exception:
            _clear_singleton()

    hosts = collect_wall_foundations_in_view(doc, active_view)
    if not hosts:
        _mostrar_aviso(
            uiapp,
            u"No hay Wall Foundations en la vista activa.",
        )
        return

    previews = []
    for wf in hosts:
        try:
            prev = build_host_preview_in_view(wf, active_view)
        except Exception:
            prev = None
        previews.append(prev)

    ctrl = FundacionCorridaFranjaController(
        uiapp, uidoc, doc, hosts, previews, active_view
    )
    try:
        ctrl.show()
    except Exception:
        _clear_singleton()
        raise
