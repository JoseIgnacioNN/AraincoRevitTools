# -*- coding: utf-8 -*-
"""
Armadura en vigas — capa superior e inferior y barras laterales (piel).

**Entorno:** Revit 2024–2026 · Python **IronPython 2.7** (pyRevit). La API puede variar entre versiones.

**Extensión BIMTools:** convenciones de UI (`bimtools-ui.mdc`: manual en carpeta del pushbutton),
transacciones y manejo de errores (`revit-api-python.mdc`).

Flujo API: ventana no modal · ExternalEvent: PickObjects (vigas) · ExternalEvent: Transaction crear Rebar.
Recubrimiento fijo 25 mm; eje en apoyos se extiende para cubrir el elemento (bbox) menos recubrimiento en la cara de salida.
Laterales: misma lógica de eje y apoyos que las capas (extensión, hosts vecinos); en cadenas colineales, dos Rebar (cara +w y −w) en el tramo fusionado, como sup/inf. La franja en canto y en ancho se calcula con bbox solo de vigas (Structural Framing), no de columnas, para no desplazar la piel fuera del alma. Recorte longitudinal con sólido del host; 100 mm libres en canto respecto a fibras sup/inf.
"""

# Recubrimiento fijo (mm) para esta herramienta.
_COVER_MM_FIXED = 25.0

# Separación adicional (mm) entre laterales y la primera capa sup/inf (además del recubrimiento).
_LATERAL_CLEAR_FROM_FLEXURAL_MM = 100.0

# Paso (mm) para proponer nº de laterales: floor(altura_mm / paso) − 1 (mín. 1).
_LATERAL_COUNT_STEP_MM = 200.0

import os
import weakref
import clr
import math

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

import System
from System.Collections.Generic import List
from System.Windows.Markup import XamlReader
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, RoutedEventHandler
from System.Windows.Controls import CheckBox, ComboBoxItem
from System.Windows.Input import Key, KeyBinding, ModifierKeys, ApplicationCommands, CommandBinding

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    GeometryInstance,
    Line,
    LocationCurve,
    Options,
    Solid,
    SolidCurveIntersectionMode,
    SolidCurveIntersectionOptions,
    Transaction,
    ViewDetailLevel,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarHookType,
    RebarShape,
    RebarStyle,
    StructuralType,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

# XAML alineado con area_reinforcement_losa (Crear Area Reinforcement RPS) — tema Arainco oscuro.
XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="BIMTools — Armadura en vigas"
        Height="900" Width="520" MinHeight="560" MinWidth="480"
        WindowStartupLocation="Manual"
        Background="#0A1C26"
        FontFamily="Segoe UI"
        ResizeMode="NoResize">
  <Window.Resources>
    <Style x:Key="Label" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#4A8BA6"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,10,0,4"/>
    </Style>
    <Style x:Key="LabelSmall" TargetType="TextBlock" BasedOn="{StaticResource Label}">
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="Margin" Value="0,4,0,2"/>
    </Style>
    <Style x:Key="Combo" TargetType="ComboBox">
      <Setter Property="Background" Value="#0D2234"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#1A3D52"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Height" Value="32"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid TextElement.Foreground="{TemplateBinding Foreground}">
              <Border x:Name="Border"
                      Background="{TemplateBinding Background}"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}"
                      CornerRadius="6"/>
              <ToggleButton IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                            Focusable="False" Background="Transparent" BorderThickness="0"
                            HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>
              <ContentPresenter x:Name="ContentSite"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                Margin="12,0,32,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
              <TextBox x:Name="PART_EditableTextBox"
                       Visibility="Collapsed"
                       Background="Transparent" Foreground="{TemplateBinding Foreground}"
                       BorderThickness="0" Margin="12,0,32,0" VerticalAlignment="Center"
                       FontSize="{TemplateBinding FontSize}" FontFamily="{TemplateBinding FontFamily}"
                       CaretBrush="#5BB8D4" Padding="0" VerticalContentAlignment="Center"/>
              <TextBlock Text="&#9660;" FontSize="9" Foreground="#5BB8D4"
                         HorizontalAlignment="Right" VerticalAlignment="Center"
                         Margin="0,0,12,0" IsHitTestVisible="False"/>
              <Popup x:Name="PART_Popup"
                     IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False"
                     PopupAnimation="Fade" Placement="Bottom">
                <Border Background="#0D2234" BorderBrush="#1A3D52" BorderThickness="1"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <ScrollViewer MaxHeight="200" VerticalScrollBarVisibility="Auto">
                    <ItemsPresenter/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEditable" Value="True">
                <Setter TargetName="ContentSite" Property="Visibility" Value="Collapsed"/>
                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsEditable" Value="False">
                <Setter TargetName="ContentSite" Property="Visibility" Value="Visible"/>
                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Collapsed"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Border" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="Border" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="ComboItem" TargetType="ComboBoxItem">
      <Setter Property="Background" Value="#0D2234"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="Padding" Value="10,8"/>
      <Style.Triggers>
        <Trigger Property="IsHighlighted" Value="True">
          <Setter Property="Background" Value="#1A4F6A"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#1A4F6A"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="TxtField" TargetType="TextBox">
      <Setter Property="Background" Value="#0D2234"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#1A3D52"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Height" Value="32"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="CaretBrush" Value="#5BB8D4"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Background" Value="#5BB8D4"/>
      <Setter Property="Foreground" Value="#0A1C26"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="20,9"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="7"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#7CCDE2"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="#4AA5C0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnGhost" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="Background" Value="#0F2535"/>
      <Setter Property="Foreground" Value="#C8E4EF"/>
    </Style>
  </Window.Resources>

  <Grid Margin="16,12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#0F2535" CornerRadius="6" Padding="12,10" Margin="0,0,0,12">
      <StackPanel VerticalAlignment="Center">
        <TextBlock Text="ARMADURA EN VIGAS" FontSize="16" FontWeight="Bold" Foreground="#C8E4EF"/>
        <TextBlock Text="Recubrimiento 25 mm fijo; capas: eje prolongado en apoyos (continuo en cadena). Laterales: solo dentro de cada viga, piel en ambas caras del alma."
                   FontSize="11" Foreground="#4A8BA6" Margin="0,4,0,0" TextWrapping="Wrap"/>
      </StackPanel>
    </Border>

    <StackPanel Grid.Row="1" Margin="0,0,0,10">
      <Button x:Name="BtnSeleccionar" Content="SELECCIONAR VIGAS EN MODELO"
              Style="{StaticResource BtnGhost}"
              HorizontalAlignment="Stretch" Padding="16,12"/>
      <TextBlock x:Name="TxtVigaInfo" Text="Ninguna viga seleccionada."
                 Foreground="#4A8BA6" FontSize="11" Margin="0,6,0,0" TextWrapping="Wrap"/>
    </StackPanel>

    <Border Grid.Row="2" Background="#0F2535" CornerRadius="6" Padding="12,10" Margin="0,0,0,10"
            BorderBrush="#1A3D52" BorderThickness="1">
      <StackPanel>
        <TextBlock Text="CAPA SUPERIOR" Style="{StaticResource Label}" Margin="0,0,0,8"/>
        <Border Background="#0A1C26" CornerRadius="4" Padding="8,6" BorderBrush="#1A3D52" BorderThickness="1">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
              <TextBlock Text="Tipo de barra" Style="{StaticResource LabelSmall}"/>
              <ComboBox x:Name="CmbSup" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                  <ComboBox.ItemContainerStyle>
                    <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
                  </ComboBox.ItemContainerStyle>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Column="2">
              <TextBlock Text="Nº barras" Style="{StaticResource LabelSmall}"/>
              <TextBox x:Name="TxtNsup" Style="{StaticResource TxtField}" Text="2"/>
            </StackPanel>
          </Grid>
        </Border>
      </StackPanel>
    </Border>

    <Border Grid.Row="3" Background="#0F2535" CornerRadius="6" Padding="12,10" Margin="0,0,0,10"
            BorderBrush="#1A3D52" BorderThickness="1">
      <StackPanel>
        <TextBlock Text="CAPA INFERIOR" Style="{StaticResource Label}" Margin="0,0,0,8"/>
        <Border Background="#0A1C26" CornerRadius="4" Padding="8,6" BorderBrush="#1A3D52" BorderThickness="1">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
              <TextBlock Text="Tipo de barra" Style="{StaticResource LabelSmall}"/>
              <ComboBox x:Name="CmbInf" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                  <ComboBox.ItemContainerStyle>
                    <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
                  </ComboBox.ItemContainerStyle>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Column="2">
              <TextBlock Text="Nº barras" Style="{StaticResource LabelSmall}"/>
              <TextBox x:Name="TxtNinf" Style="{StaticResource TxtField}" Text="2"/>
            </StackPanel>
          </Grid>
        </Border>
      </StackPanel>
    </Border>

    <Border Grid.Row="4" Background="#0F2535" CornerRadius="6" Padding="12,10" Margin="0,0,0,10"
            BorderBrush="#1A3D52" BorderThickness="1">
      <StackPanel>
        <TextBlock Text="BARRAS LATERALES (PIEL)" Style="{StaticResource Label}" Margin="0,0,0,8"/>
        <CheckBox x:Name="ChkLaterales" Content="Incluir barras en caras laterales del alma (reparto en altura)"
                  Foreground="#C8E4EF" FontSize="11" Margin="0,0,0,8"/>
        <Border Background="#0A1C26" CornerRadius="4" Padding="8,6" BorderBrush="#1A3D52" BorderThickness="1">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
              <TextBlock Text="Tipo de barra" Style="{StaticResource LabelSmall}"/>
              <ComboBox x:Name="CmbLat" Style="{StaticResource Combo}" IsEditable="False" IsReadOnly="True">
                  <ComboBox.ItemContainerStyle>
                    <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
                  </ComboBox.ItemContainerStyle>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Column="2">
              <TextBlock Text="Nº barras / cara (def.: altura÷200−1)" Style="{StaticResource LabelSmall}"/>
              <TextBox x:Name="TxtNlat" Style="{StaticResource TxtField}" Text="1"/>
            </StackPanel>
          </Grid>
        </Border>
      </StackPanel>
    </Border>

    <StackPanel Grid.Row="5" Margin="0,0,0,10">
      <TextBlock x:Name="TxtEstado" Text="" Foreground="#5BB8D4" FontSize="11"
                 Margin="0,0,0,8" TextWrapping="Wrap"/>
      <Button x:Name="BtnColocar" Content="COLOCAR ARMADURAS"
              Style="{StaticResource BtnPrimary}"
              HorizontalAlignment="Stretch" Padding="20,12"/>
    </StackPanel>

    <Grid Grid.Row="6" Margin="0,4,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <Button x:Name="BtnManual" Grid.Column="0" Content="MANUAL"
              Style="{StaticResource BtnGhost}" ToolTip="Abrir manual de usuario"
              Padding="16,10" Margin="0,0,8,0" MinWidth="88"/>
      <Button x:Name="BtnCancel" Grid.Column="2" Content="CERRAR"
              Style="{StaticResource BtnGhost}" MinWidth="100" Padding="16,10"/>
    </Grid>

    <Border Grid.Row="7" Background="Transparent"/>
  </Grid>
</Window>
"""


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _ray_aabb_interval_positive(origin, dir_u, mn, mx):
    """
    Intersección del rayo origin + t*dir_u (dir_u no necesariamente unitario; t>=0)
    con el paralelepípedo alineado a ejos [mn,mx]. Retorna (t_enter, t_exit) o None.
    """
    try:
        ox, oy, oz = float(origin.X), float(origin.Y), float(origin.Z)
        dx, dy, dz = float(dir_u.X), float(dir_u.Y), float(dir_u.Z)
        xmin, ymin, zmin = float(mn.X), float(mn.Y), float(mn.Z)
        xmax, ymax, zmax = float(mx.X), float(mx.Y), float(mx.Z)
    except Exception:
        return None
    tmin = 0.0
    tmax = 1.0e30
    pairs = ((ox, dx, xmin, xmax), (oy, dy, ymin, ymax), (oz, dz, zmin, zmax))
    for oc, dc, bmn, bmx in pairs:
        if abs(dc) < 1e-12:
            if oc < bmn - 1e-9 or oc > bmx + 1e-9:
                return None
            continue
        inv = 1.0 / dc
        t1 = (bmn - oc) * inv
        t2 = (bmx - oc) * inv
        if t1 > t2:
            t1, t2 = t2, t1
        tmin = max(tmin, t1)
        tmax = min(tmax, t2)
        if tmin > tmax + 1e-9:
            return None
    if tmax < 0:
        return None
    return (max(0.0, tmin), tmax)


def _best_extension_into_support_ft(document, p_end, dir_into, cover_ft, exclude_ids):
    """
    Cuánto alargar el eje desde p_end hacia dir_into (unitario) hasta cubrir el apoyo
    (columna o viga) según su bbox: profundidad del rayo menos 1× recubrimiento en la cara de salida.
    Prioriza columnas sobre vigas de pórtico. Sin tope máximo de penetración.
    Retorna (extensión_ft, elemento_o_None).
    """
    exclude_ids = set(exclude_ids or [])
    best_ext = 0.0
    best_el = None

    for e in FilteredElementCollector(document).OfCategory(
        BuiltInCategory.OST_StructuralColumns
    ).WhereElementIsNotElementType():
        if not isinstance(e, FamilyInstance):
            continue
        try:
            if int(e.Id.IntegerValue) in exclude_ids:
                continue
        except Exception:
            continue
        bb = e.get_BoundingBox(None)
        if bb is None:
            continue
        iv = _ray_aabb_interval_positive(p_end, dir_into, bb.Min, bb.Max)
        if iv is None:
            continue
        t0, t1 = iv
        depth_in = float(t1) - max(0.0, float(t0))
        if depth_in < 1e-5:
            continue
        ext = max(0.0, depth_in - 1.0 * float(cover_ft))
        if ext > best_ext + 1e-7:
            best_ext = ext
            best_el = e

    if best_ext < 1e-7:
        for e in FilteredElementCollector(document).OfCategory(
            BuiltInCategory.OST_StructuralFraming
        ).WhereElementIsNotElementType():
            if not isinstance(e, FamilyInstance):
                continue
            try:
                if getattr(e, "StructuralType", None) != StructuralType.Beam:
                    continue
            except Exception:
                continue
            try:
                if int(e.Id.IntegerValue) in exclude_ids:
                    continue
            except Exception:
                continue
            bb = e.get_BoundingBox(None)
            if bb is None:
                continue
            iv = _ray_aabb_interval_positive(p_end, dir_into, bb.Min, bb.Max)
            if iv is None:
                continue
            t0, t1 = iv
            depth_in = float(t1) - max(0.0, float(t0))
            if depth_in < 1e-5:
                continue
            ext = max(0.0, depth_in - 1.0 * float(cover_ft))
            if ext > best_ext + 1e-7:
                best_ext = ext
                best_el = e
    return best_ext, best_el


def _extend_span_to_adjacent_supports(
    document,
    beam_elem,
    p0,
    p1,
    axis_unit,
    cover_ft,
    extra_exclude_ids=None,
):
    """
    Alarga p0/p1 hacia el interior de columnas o vigas colindantes en los apoyos.
    axis_unit: de p0 hacia p1 (normalizado).
    Retorna (p0_nueva, p1_nueva, vecinos, ext0_ft, ext1_ft) para no duplicar recubrimiento
    en eje si ya se penetró en el apoyo (evita ~3× rec en cara de columna).
    """
    ex = set(extra_exclude_ids or [])
    try:
        ex.add(int(beam_elem.Id.IntegerValue))
    except Exception:
        pass
    ext0, el0 = _best_extension_into_support_ft(
        document, p0, axis_unit.Negate(), cover_ft, ex
    )
    ext1, el1 = _best_extension_into_support_ft(
        document, p1, axis_unit, cover_ft, ex
    )
    p0n = p0 - axis_unit * ext0
    p1n = p1 + axis_unit * ext1
    neigh = []
    if el0 is not None:
        neigh.append(el0)
    if el1 is not None and (el0 is None or int(el1.Id.IntegerValue) != int(el0.Id.IntegerValue)):
        neigh.append(el1)
    return p0n, p1n, neigh, float(ext0), float(ext1)


def _dedupe_elements_by_id(elems):
    seen = set()
    out = []
    for e in elems:
        if e is None:
            continue
        try:
            k = int(e.Id.IntegerValue)
        except Exception:
            continue
        if k in seen:
            continue
        seen.add(k)
        out.append(e)
    return out


def _collect_bar_types(document):
    rows = []
    seen = set()
    for bt in FilteredElementCollector(document).OfClass(RebarBarType):
        if bt is None:
            continue
        try:
            eid = int(bt.Id.IntegerValue)
            if eid in seen:
                continue
            seen.add(eid)
            diam_ft = float(bt.BarNominalDiameter)
            if diam_ft <= 0:
                diam_ft = float(getattr(bt, "BarModelDiameter", 0) or 0)
            diam_mm = int(round(diam_ft * 304.8)) if diam_ft > 0 else 0
            label = u"\u00f8{} mm".format(diam_mm) if diam_mm > 0 else (bt.Name or u"Barra")
            rows.append((label, bt.Id))
        except Exception:
            continue
    rows.sort(key=lambda x: x[0])
    return rows


def _pick_hook_90(document):
    pi_2 = math.pi / 2.0

    def es_90(ht):
        try:
            if getattr(ht, "Style", None) != RebarStyle.Standard:
                return False
            ang = getattr(ht, "HookAngle", None)
            return ang is not None and abs(float(ang) - pi_2) < 0.1
        except Exception:
            return False

    hook_type = None
    for ht in FilteredElementCollector(document).OfClass(RebarHookType):
        try:
            if not ht or not es_90(ht):
                continue
            name = (getattr(ht, "Name", None) or u"").strip().lower()
            if any(k in name for k in ["90", "hook", "gancho", "deg"]):
                hook_type = ht
                break
        except Exception:
            continue
    if not hook_type:
        for ht in FilteredElementCollector(document).OfClass(RebarHookType):
            if es_90(ht):
                hook_type = ht
                break
    return hook_type


def _pick_rebar_shapes(document):
    simple, others = [], []
    for shape in FilteredElementCollector(document).OfClass(RebarShape):
        try:
            if shape.RebarStyle != RebarStyle.Standard:
                continue
            if getattr(shape, "SimpleLine", False):
                simple.append(shape)
            else:
                others.append(shape)
        except Exception:
            continue
    return simple if simple else others


class VigasArmaduraFilter(ISelectionFilter):
    """Solo vigas (Structural Framing con StructuralType.Beam)."""

    _cat = int(BuiltInCategory.OST_StructuralFraming)

    def AllowElement(self, elem):
        try:
            if elem is None or elem.Category is None:
                return False
            if int(elem.Category.Id.IntegerValue) != self._cat:
                return False
            if not isinstance(elem, FamilyInstance):
                return False
            st = getattr(elem, "StructuralType", None)
            return st == StructuralType.Beam
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


def _read_width_depth_ft(document, elem, curve):
    """Ancho (luz entre caras laterales) y canto (altura de sección) en pies internos."""
    et = document.GetElement(elem.GetTypeId()) if elem else None
    w, d = None, None
    if et:
        for n in ("Width", "Ancho", "Ancho nominal", "b", "B"):
            p = et.LookupParameter(n)
            if p and p.HasValue:
                w = float(p.AsDouble())
                break
        for n in ("Height", "Depth", "Altura", "Profundidad", "h", "H", "d"):
            p = et.LookupParameter(n)
            if p and p.HasValue:
                d = float(p.AsDouble())
                break
    bb = elem.get_BoundingBox(None)
    if bb is not None:
        dx = abs(bb.Max.X - bb.Min.X)
        dy = abs(bb.Max.Y - bb.Min.Y)
        dz = abs(bb.Max.Z - bb.Min.Z)
        dims = sorted([dx, dy, dz], reverse=True)
        small = sorted(dims[1:]) if len(dims) >= 3 else [dims[-1], dims[-1]]
        bbox_w = float(small[0])
        bbox_d = float(small[1])
        if not w or w <= 0:
            w = bbox_w
        if not d or d <= 0:
            d = bbox_d
    if not w or w <= 0:
        w = 1.0
    if not d or d <= 0:
        d = 1.0
    return w, d


def _default_n_lat_from_depth_mm(depth_mm):
    """Nº sugerido de laterales por cara: floor(altura_mm / 200) − 1; mínimo 1."""
    if depth_mm <= 0:
        return 1
    step = float(_LATERAL_COUNT_STEP_MM)
    n = int(math.floor(float(depth_mm) / step)) - 1
    return max(1, n)


def _default_n_lat_from_beams(document, beam_ids):
    """Mayor canto (altura de sección) entre las vigas; convierte a mm y aplica la regla."""
    max_d_mm = 0.0
    for bid in beam_ids:
        try:
            elem = document.GetElement(bid)
        except Exception:
            elem = None
        if elem is None:
            continue
        loc = getattr(elem, "Location", None)
        if not isinstance(loc, LocationCurve):
            continue
        try:
            curve = loc.Curve
            _, d_ft = _read_width_depth_ft(document, elem, curve)
            d_mm = float(d_ft) * 304.8
            if d_mm > max_d_mm:
                max_d_mm = d_mm
        except Exception:
            continue
    return _default_n_lat_from_depth_mm(max_d_mm)


def _beam_frame(curve):
    """
    Eje de la viga (unitario), dirección de ancho (horizontal, ⟂ eje) y dirección de canto (hacia arriba típico).
    """
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    raw = p1 - p0
    ln = raw.GetLength()
    if ln < 1e-9:
        return None
    axis = raw.Normalize()
    if abs(axis.Z) > 0.92:
        return None
    z_up = XYZ.BasisZ
    width_dir = axis.CrossProduct(z_up)
    if width_dir.GetLength() < 1e-9:
        width_dir = axis.CrossProduct(XYZ.BasisX)
    width_dir = width_dir.Normalize()
    depth_dir = width_dir.CrossProduct(axis).Normalize()
    if depth_dir.Z < 0:
        depth_dir = depth_dir.Negate()
        width_dir = axis.CrossProduct(depth_dir).Normalize()
    return axis, width_dir, depth_dir, p0, p1, ln


def _beam_curve_length_ft(elem):
    try:
        c = elem.Location.Curve
        return float(c.GetEndPoint(0).DistanceTo(c.GetEndPoint(1)))
    except Exception:
        return 0.0


def _beam_placement_data(document, elem):
    """Datos geométricos para agrupar vigas colineales; None si no aplica."""
    loc = elem.Location
    if not isinstance(loc, LocationCurve):
        return None
    curve = loc.Curve
    frame = _beam_frame(curve)
    if not frame:
        return None
    axis, width_dir, depth_dir, p0, p1, beam_len = frame
    return {
        "elem": elem,
        "curve": curve,
        "axis": axis,
        "width_dir": width_dir,
        "depth_dir": depth_dir,
        "p0": p0,
        "p1": p1,
        "beam_len": beam_len,
    }


def _dist_point_to_ray_ft(pt, origin, unit_dir):
    v = pt - origin
    t = float(v.DotProduct(unit_dir))
    perp = v - unit_dir * t
    return float(perp.GetLength())


def _beams_collinear_for_cluster(di, dj, tol_perp_ft):
    """Misma recta de eje (tolerancia perpendicular), ejes paralelos."""
    axis_i = di["axis"].Normalize()
    axis_j = dj["axis"].Normalize()
    if abs(float(axis_i.DotProduct(axis_j))) < 0.98:
        return False
    p0i, p1i = di["p0"], di["p1"]
    p0j, p1j = dj["p0"], dj["p1"]
    if _dist_point_to_ray_ft(p0j, p0i, axis_i) > tol_perp_ft:
        return False
    if _dist_point_to_ray_ft(p1j, p0i, axis_i) > tol_perp_ft:
        return False
    if _dist_point_to_ray_ft(p0i, p0j, axis_j) > tol_perp_ft:
        return False
    if _dist_point_to_ray_ft(p1i, p0j, axis_j) > tol_perp_ft:
        return False
    return True


def _cluster_beam_indices(beam_data_list, tol_perp_ft=0.5):
    """Índices agrupados en cadenas colineales (~15 cm por defecto si 0.5 ft)."""
    n = len(beam_data_list)
    parent = list(range(n))
    rank = [0] * n

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        ri, rj = find(i), find(j)
        if ri == rj:
            return
        if rank[ri] < rank[rj]:
            parent[ri] = rj
        elif rank[ri] > rank[rj]:
            parent[rj] = ri
        else:
            parent[rj] = ri
            rank[ri] += 1

    for i in range(n):
        for j in range(i + 1, n):
            if _beams_collinear_for_cluster(beam_data_list[i], beam_data_list[j], tol_perp_ft):
                union(i, j)
    clusters = {}
    for i in range(n):
        r = find(i)
        clusters.setdefault(r, []).append(i)
    return clusters.values()


def _order_cluster_indices(beam_data_list, indices, axis, p0_ref):
    axis = axis.Normalize()

    def mid(ii):
        d = beam_data_list[ii]
        return 0.5 * (d["p0"] + d["p1"])

    return sorted(indices, key=lambda ii: float((mid(ii) - p0_ref).DotProduct(axis)))


def _union_bbox_corners_from_elems(elems):
    corners = []
    for e in elems:
        if e is None:
            continue
        try:
            b = e.get_BoundingBox(None)
        except Exception:
            continue
        if b is None:
            continue
        corners.extend(_bbox_corners_xyz(b))
    return corners


def _rebar_ortho_frame_t_w_n(axis, width_dir):
    """
    Triedro ortonormal: t = eje viga, w = ancho (⟂ t), n = t × w (canto, Z modelo preferente).

    En Revit, el reparto de Fixed Number se extiende a lo largo de la NORMAL del plano de la
    barra (o el lado opuesto con barsOnNormalSide), no en tangente×normal. Por eso:
    - normal = n (canto) → barras se apilan en vertical en sección (mal para capas sup/inf).
    - normal = w (ancho) → el set se abre en el ancho de la viga (lo deseado).
    La curva yace en el plano ⟂ w (contiene t y n).
    """
    if axis is None or width_dir is None:
        return None
    try:
        t = axis.Normalize()
        w = width_dir - t * float(t.DotProduct(width_dir))
        if w.GetLength() < 1e-9:
            w = t.CrossProduct(XYZ.BasisZ)
            if w.GetLength() < 1e-9:
                w = t.CrossProduct(XYZ.BasisX)
        w = w.Normalize()
        n = t.CrossProduct(w).Normalize()
        if float(n.Z) < 0.0:
            n = n.Negate()
            w = n.CrossProduct(t).Normalize()
        return t, w, n
    except Exception:
        return None


def _shift_line_mid_along_width_to_bbox_center(elem, p0, p1, width_dir, width_ft):
    """
    Acerca la línea guía al centro del bbox en dirección ancho. En vigas en T el centro
    del bbox cae hacia el alma: un desplazamiento grande empuja un barra del set fuera del ala.
    Se limita |s| a una fracción del ancho de tipo (p. ej. 30 %).
    """
    bb = elem.get_BoundingBox(None)
    if bb is None or width_dir is None or width_dir.GetLength() < 1e-9:
        return XYZ(0, 0, 0)
    try:
        mid = p0 + (p1 - p0) * 0.5
        c = XYZ(
            0.5 * (float(bb.Min.X) + float(bb.Max.X)),
            0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
            0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
        )
        d = c - mid
        s = float(d.DotProduct(width_dir))
        lim = max(0.0, 0.30 * float(width_ft))
        if lim > 1e-9:
            s = max(-lim, min(lim, s))
        return width_dir * s
    except Exception:
        return XYZ(0, 0, 0)


def _shift_line_mid_along_width_to_bbox_union(elems, p0, p1, width_dir, width_ft):
    """Igual que _shift_line_mid_along_width_to_bbox_center pero bbox unión de varias vigas."""
    hc = _union_bbox_corners_from_elems(elems)
    if not hc or width_dir is None or width_dir.GetLength() < 1e-9:
        return XYZ(0, 0, 0)
    try:
        mid = p0 + (p1 - p0) * 0.5
        xs = [float(p.X) for p in hc]
        ys = [float(p.Y) for p in hc]
        zs = [float(p.Z) for p in hc]
        c = XYZ(0.5 * (min(xs) + max(xs)), 0.5 * (min(ys) + max(ys)), 0.5 * (min(zs) + max(zs)))
        d = c - mid
        s = float(d.DotProduct(width_dir))
        lim = max(0.0, 0.30 * float(width_ft))
        if lim > 1e-9:
            s = max(-lim, min(lim, s))
        return width_dir * s
    except Exception:
        return XYZ(0, 0, 0)


def _layer_dz_world_z_from_bbox(elem, p0, p1, cov_ft):
    """
    Desplazamiento en Z modelo respecto al punto medio de la LocationCurve para
    situar capas en la fibra interior (cara inf / sup) según el BoundingBox real.
    Evita asumir que la curva pasa por el centro de sección (Revit: ref. arriba/abajo/centro).
    Retorna (dz_inf, dz_sup) o None si no aplica.
    """
    bb = elem.get_BoundingBox(None)
    if bb is None:
        return None
    z_span = float(bb.Max.Z - bb.Min.Z)
    if z_span < 2.0 * cov_ft + 1e-6:
        return None
    z_mid_curve = 0.5 * (float(p0.Z) + float(p1.Z))
    z_bot_inner = float(bb.Min.Z) + cov_ft
    z_top_inner = float(bb.Max.Z) - cov_ft
    if z_top_inner <= z_bot_inner + 1e-6:
        return None
    dz_inf = z_bot_inner - z_mid_curve
    dz_sup = z_top_inner - z_mid_curve
    return dz_inf, dz_sup


def _layer_dz_world_z_from_bbox_multi(elems, p0, p1, cov_ft):
    """Capas sup/inf respecto a la envolvente Z de varias vigas colineales."""
    hc = _union_bbox_corners_from_elems(elems)
    if not hc:
        return None
    try:
        zs = [float(p.Z) for p in hc]
        z_min = min(zs)
        z_max = max(zs)
        z_span = z_max - z_min
        if z_span < 2.0 * cov_ft + 1e-6:
            return None
        z_mid_curve = 0.5 * (float(p0.Z) + float(p1.Z))
        z_bot_inner = z_min + cov_ft
        z_top_inner = z_max - cov_ft
        if z_top_inner <= z_bot_inner + 1e-6:
            return None
        dz_inf = z_bot_inner - z_mid_curve
        dz_sup = z_top_inner - z_mid_curve
        return dz_inf, dz_sup
    except Exception:
        return None


def _invert_hook_orientation(o):
    """Intercambia Left/Right para voltear ganchos hacia el interior en capa inferior."""
    if o == RebarHookOrientation.Right:
        return RebarHookOrientation.Left
    return RebarHookOrientation.Right


def _bbox_corners_xyz(bb):
    """8 esquinas del BoundingBox (coordenadas modelo)."""
    if bb is None:
        return []
    mn, mx = bb.Min, bb.Max
    out = []
    for x in (float(mn.X), float(mx.X)):
        for y in (float(mn.Y), float(mx.Y)):
            for z in (float(mn.Z), float(mx.Z)):
                out.append(XYZ(x, y, z))
    return out


def _collect_solids_from_host_element(elem):
    """Sólidos de get_Geometry del host (instancia o directo)."""
    if elem is None:
        return []
    opts = Options()
    opts.ComputeReferences = False
    opts.IncludeNonVisibleObjects = False
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        ge = elem.get_Geometry(opts)
    except Exception:
        return []
    if ge is None:
        return []
    out = []
    for obj in ge:
        if obj is None:
            continue
        if isinstance(obj, Solid) and obj.Volume > 1e-12:
            out.append(obj)
        elif isinstance(obj, GeometryInstance):
            try:
                sub = obj.GetInstanceGeometry()
                if sub is not None:
                    for g in sub:
                        if isinstance(g, Solid) and g.Volume > 1e-12:
                            out.append(g)
            except Exception:
                pass
    return out


def _merge_axis_intervals(intervals):
    """Une intervalos [a,b] en R (pies a lo largo del parámetro de curva)."""
    if not intervals:
        return []
    iv = sorted(intervals, key=lambda x: x[0])
    merged = []
    for a, b in iv:
        if b < a:
            a, b = b, a
        if not merged or a > merged[-1][1] + 1e-9:
            merged.append([a, b])
        else:
            merged[-1][1] = max(merged[-1][1], b)
    return merged


def _solid_line_inside_param_intervals(line, solid):
    """Tramos del segmento que caen dentro del sólido (parámetros nativos de la curva)."""
    out = []
    try:
        scio = SolidCurveIntersectionOptions()
        try:
            scio.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside
        except Exception:
            pass
        sci = solid.IntersectWithCurve(line, scio)
    except Exception:
        return out
    if sci is None:
        return out
    try:
        n = int(sci.SegmentCount)
    except Exception:
        return out
    if n < 1:
        return out
    for i in range(n):
        try:
            ext = sci.GetCurveSegmentExtents(i)
            s0 = float(ext.StartParameter)
            s1 = float(ext.EndParameter)
            if s1 < s0:
                s0, s1 = s1, s0
            out.append((s0, s1))
        except Exception:
            continue
    return out


def _best_merged_span_for_midpoint(merged, curve_len, mid_param):
    """Elige el intervalo que contiene el punto medio del tramo analítico, o el de mayor solape."""
    if not merged:
        return None
    for a, b in merged:
        if a <= mid_param <= b:
            return (float(a), float(b))
    best = None
    best_len = -1.0
    for a, b in merged:
        lo = max(0.0, float(a))
        hi = min(float(curve_len), float(b))
        span = hi - lo
        if span > best_len:
            best_len = span
            best = (float(a), float(b))
    return best if best_len > 1e-9 else None


def _axis_solid_span_params(p0, p1, lateral_offset, solids):
    """
    Parámetros a lo largo de la línea A=p0+off → B=p1+off donde el hormigón envuelve el trazo.
    Los parámetros coinciden con la distancia desde A si la curva es una línea recta acotada.
    Retorna (t0, t1) en [0, L] respecto al mismo eje que p0→p1, o None.
    """
    if not solids:
        return None
    try:
        off = lateral_offset if lateral_offset is not None else XYZ(0, 0, 0)
        a = p0 + off
        b = p1 + off
        ln_seg = Line.CreateBound(a, b)
        clen = float(ln_seg.Length)
    except Exception:
        return None
    if clen < 1e-9:
        return None
    acc = []
    for solid in solids:
        acc.extend(_solid_line_inside_param_intervals(ln_seg, solid))
    merged = _merge_axis_intervals(acc)
    mid = 0.5 * clen
    span = _best_merged_span_for_midpoint(merged, clen, mid)
    if span is None:
        return None
    s0, s1 = span
    if s1 <= s0 + 1e-9:
        return None
    return (max(0.0, s0), min(clen, s1))


def _axis_bbox_t_span_clamped(p0, p1, host, margin):
    """
    (t0, t1) desde p0 a lo largo del eje p0→p1 por proyección del bbox del host, con márgenes.
    """
    raw = p1 - p0
    ln = float(raw.GetLength())
    if ln < 1e-9:
        return None
    ax = raw.Normalize()
    m = float(margin)
    bb = host.get_BoundingBox(None) if host else None
    if bb is None:
        return (m, ln - m)
    hc = _bbox_corners_xyz(bb)
    if not hc:
        return (m, ln - m)
    ts = [float((c - p0).DotProduct(ax)) for c in hc]
    t_min = min(ts)
    t_max = max(ts)
    t_lo = t_min + m
    t_hi = t_max - m
    t_u0 = max(t_lo, 0.0, m)
    t_u1 = min(t_hi, ln, ln - m)
    if t_u1 <= t_u0 + 1e-6:
        return None
    return (t_u0, t_u1)


def _axis_cover_trim_endpoints(
    p0, p1, host, cov_ft, bar_diam_ft=0.0, lateral_offset_xyz=None
):
    """
    Eje de barra recortado al hormigón del host: proyección del bbox sobre p0→p1,
    menos (recubrimiento + radio nominal) en cada extremo longitudinal respecto a la cara;
    intersección con el tramo analítico [0, L].

    Si lateral_offset_xyz no es None, se cruza además con el tramo real del sólido del host
    a lo largo de la línea desplazada a la cara lateral (evita alargar por bbox AABB o curva
    que no coincide con el canto de la viga frente al pilar).
    """
    if host is None:
        return None, None
    raw = p1 - p0
    ln = float(raw.GetLength())
    if ln < 1e-9:
        return None, None
    ax = raw.Normalize()
    margin = float(cov_ft) + 0.5 * max(float(bar_diam_ft), 1e-6)

    bb_span = _axis_bbox_t_span_clamped(p0, p1, host, margin)

    sol_span = None
    if lateral_offset_xyz is not None:
        solids = _collect_solids_from_host_element(host)
        sol_span = _axis_solid_span_params(p0, p1, lateral_offset_xyz, solids)
        if sol_span:
            s0, s1 = sol_span
            s0 = max(s0, margin)
            s1 = min(s1, ln - margin)
            if s1 > s0 + 1e-6:
                sol_span = (s0, s1)
            else:
                sol_span = None

    t_u0 = t_u1 = None
    if bb_span and sol_span:
        t_u0 = max(bb_span[0], sol_span[0])
        t_u1 = min(bb_span[1], sol_span[1])
        if t_u1 <= t_u0 + 1e-6:
            t_u0, t_u1 = sol_span[0], sol_span[1]
    elif sol_span:
        t_u0, t_u1 = sol_span[0], sol_span[1]
    elif bb_span:
        t_u0, t_u1 = bb_span[0], bb_span[1]

    if t_u0 is not None and t_u1 is not None and t_u1 > t_u0 + 1e-6:
        return p0 + ax * t_u0, p0 + ax * t_u1

    pa = p0 + ax * margin
    pb = p1 - ax * margin
    if float((pb - pa).DotProduct(ax)) < 2.0 * margin + 1e-6:
        return None, None
    return pa, pb


def _rebar_shapes_simple_line_only(shapes):
    """Formas rectas (sin dobleces de catálogo) para fallback de laterales — evita ganchos extra."""
    if not shapes:
        return []
    return [s for s in shapes if getattr(s, "SimpleLine", False)]


def _hook_bbox_extrema_along_n(rebar, host, n_section, host_elems_union=None):
    """
    Proyecciones del host y del rebar sobre n_section (canto, hacia cara superior = +).
    Retorna (hmin, hmax, rmin, rmax) en pies internos o None.
    host_elems_union: si se pasa, el envolvente del hormigón es la unión de bbox de esas vigas.
    """
    if rebar is None or n_section is None:
        return None
    try:
        n = n_section.Normalize()
    except Exception:
        return None
    if n.GetLength() < 1e-9:
        return None
    if host_elems_union:
        hc = _union_bbox_corners_from_elems(host_elems_union)
    else:
        if host is None:
            return None
        hbb = host.get_BoundingBox(None)
        if hbb is None:
            return None
        hc = _bbox_corners_xyz(hbb)
    rbb = rebar.get_BoundingBox(None)
    if rbb is None:
        return None
    rc = _bbox_corners_xyz(rbb)
    if not hc or not rc:
        return None
    hmax = max(float(p.DotProduct(n)) for p in hc)
    hmin = min(float(p.DotProduct(n)) for p in hc)
    rmax = max(float(p.DotProduct(n)) for p in rc)
    rmin = min(float(p.DotProduct(n)) for p in rc)
    return (hmin, hmax, rmin, rmax)


def _hook_stickout_penalty_ft(
    rebar, host, n_section, es_capa_inferior, host_elems_union=None
):
    """
    Cuánto sobresale la armadura del host en dirección de canto n_section (+n = hacia la cara superior).
    Capa superior: penaliza extensión por encima del host (mala = ganchos hacia arriba).
    Capa inferior: penaliza extensión por debajo del host (mala = ganchos hacia abajo).
    """
    d = _hook_bbox_extrema_along_n(rebar, host, n_section, host_elems_union)
    if d is None:
        return 1e9
    hmin, hmax, rmin, rmax = d
    tol = 0.02  # ~6 mm; ignora rozamiento numérico
    if es_capa_inferior:
        return max(0.0, float(hmin) - float(rmin) - tol)
    return max(0.0, float(rmax) - float(hmax) - tol)


def _is_rl_lr_hook_pair(pair):
    return pair in (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    )


def _hook_orientation_rank_key(
    rebar, host, n_section, es_capa_inferior, pair, host_elems_union=None
):
    """
    Clave para minimizar (mejor primero):
    - Inferior: menos salida por -n, luego rmin lo más alto posible (ganchos hacia el alma).
    - Superior: menos salida por +n, luego rmax lo más bajo posible (ganchos hacia abajo, no fuera del ala).
    En empates finales se prefiere RL/LR frente a LL/RR.
    """
    d = _hook_bbox_extrema_along_n(rebar, host, n_section, host_elems_union)
    if d is None:
        return None
    hmin, hmax, rmin, rmax = d
    tol = 0.02
    pen_inf = max(0.0, float(hmin) - float(rmin) - tol)
    pen_sup = max(0.0, float(rmax) - float(hmax) - tol)
    t_rl = 0 if _is_rl_lr_hook_pair(pair) else 1
    if es_capa_inferior:
        return (pen_inf, -float(rmin), t_rl)
    return (pen_sup, float(rmax), t_rl)


def _optimize_hook_orientations(
    rebar, host, n_section, es_capa_inferior, host_elems_union=None
):
    """
    Tras crear el rebar set, ajusta Left/Right en ambos extremos para meter los ganchos en el host.
    El primer intento de CreateFromCurvesAndShape puede dejar ganchos hacia afuera; esto no cambia el layout.

    Nota API Revit: SetHookOrientation vive en Rebar, no en RebarShapeDrivenAccessor (si se llama al
    accessor, falla en silencio dentro del try/except y los ganchos no se corrigen).

    Con solo la penalización de salida del bbox, la capa superior suele empatar en 0 entre varias
    orientaciones y quedaba la primera del bucle (a menudo mala). Se añade rmax/rmin y preferencia RL/LR.
    """
    if rebar is None or n_section is None:
        return
    if host is None and not host_elems_union:
        return
    opts = (
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
    )

    best_key = None
    best_pair = None
    for o0, o1 in opts:
        pair = (o0, o1)
        try:
            rebar.SetHookOrientation(0, o0)
            rebar.SetHookOrientation(1, o1)
            key = _hook_orientation_rank_key(
                rebar, host, n_section, es_capa_inferior, pair, host_elems_union
            )
        except Exception:
            continue
        if key is None:
            continue
        if best_key is None or key < best_key:
            best_key = key
            best_pair = pair
    if best_pair is not None:
        try:
            rebar.SetHookOrientation(0, best_pair[0])
            rebar.SetHookOrientation(1, best_pair[1])
        except Exception:
            pass


def _hook_orientation_pairs_create(es_capa_inferior):
    """
    Pares (inicio, fin) para CreateFromCurvesAndShape.
    Con reparto del set en ancho (normal = ancho), RR/LL suelen abrir ambos ganchos 90°
    hacia el exterior en sección; RL/LR alternan y suelen orientarlos hacia el interior.
    Capa inferior: primero variantes invertidas (ganchos hacia arriba al alma).
    """
    rl_lr = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    )
    same = (
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    primary = rl_lr + same
    inverted = tuple(
        (_invert_hook_orientation(a), _invert_hook_orientation(b)) for a, b in primary
    )
    if es_capa_inferior:
        return inverted + primary
    return primary + inverted


def _width_half_span_ft(width_ft, cover_ft, bar_diam_ft):
    """Mitad de la luz útil en ancho (desde eje a fibra interior), en pies."""
    margin = cover_ft + 0.5 * max(bar_diam_ft, 1e-6)
    half = 0.5 * float(width_ft) - margin
    if half <= 0:
        half = max(0.02 * float(width_ft), 1e-4)
    return half


def _depth_half_span_ft(depth_ft, cover_ft, bar_diam_ft):
    """Mitad de la luz útil en canto (desde eje de capas a fibra interior en altura), en pies."""
    margin = cover_ft + 0.5 * max(bar_diam_ft, 1e-6)
    half = 0.5 * float(depth_ft) - margin
    if half <= 0:
        half = max(0.02 * float(depth_ft), 1e-4)
    return half


def _beam_elems_for_lateral_bbox(elems, host):
    """
    Solo Structural Framing para unir bbox en w y canto (n).
    Si la lista incluye columnas de apoyo, la unión abarca toda la altura del pilar y
    _depth_span_along_n_from_bbox / _lateral_w_scalar_from_bbox sitúan las piel fuera de la viga.
    """
    beam_cat = int(BuiltInCategory.OST_StructuralFraming)
    if not elems:
        return [host] if host is not None else []
    out = []
    for e in elems:
        try:
            if e is None:
                continue
            cat = e.Category
            if cat is None:
                continue
            if int(cat.Id.IntegerValue) != beam_cat:
                continue
            out.append(e)
        except Exception:
            continue
    out = _dedupe_elements_by_id(out)
    if out:
        return out
    return [host] if host is not None else []


def _depth_span_along_n_from_bbox(
    host_elems,
    curve_mid,
    n_plane,
    cov_ft,
    bar_diam_ft,
    extra_from_bottom_fiber_ft=0.0,
    extra_from_top_fiber_ft=0.0,
):
    """
    Centro de la franja útil y mitad de luz en dirección canto (n) proyectando el bbox del hormigón.
    extra_from_*: hueco adicional hacia el interior desde fibra inf/sup (p. ej. no solapar capas flexión).
    Retorna (n_strip_mid, usable_half) o None.
    """
    if not host_elems or n_plane is None:
        return None
    try:
        n_unit = n_plane.Normalize()
    except Exception:
        return None
    if n_unit.GetLength() < 1e-9:
        return None
    hc = _union_bbox_corners_from_elems(host_elems)
    if not hc:
        return None
    projs = []
    for c in hc:
        projs.append(float((c - curve_mid).DotProduct(n_unit)))
    n_min = min(projs)
    n_max = max(projs)
    margin = cov_ft + 0.5 * max(bar_diam_ft, 1e-6)
    eb = float(extra_from_bottom_fiber_ft)
    et = float(extra_from_top_fiber_ft)
    n_bot = n_min + margin + eb
    n_top = n_max - margin - et
    if n_top <= n_bot + 1e-6:
        return None
    n_strip_mid = 0.5 * (n_bot + n_top)
    usable_half = 0.5 * (n_top - n_bot)
    return n_strip_mid, usable_half


def _lateral_w_scalar_from_bbox(host_elems, curve_mid, w_lay, w_face_sign, cov_ft, bar_diam_ft):
    """
    Proyección escalar en w_lay (desde curve_mid) del eje de barra en la cara lateral interior.
    Evita usar ±half_w respecto al eje analítico cuando éste no coincide con el centro de alma.
    """
    if not host_elems or w_lay is None:
        return None
    try:
        w_unit = w_lay.Normalize()
    except Exception:
        return None
    if w_unit.GetLength() < 1e-9:
        return None
    hc = _union_bbox_corners_from_elems(host_elems)
    if not hc:
        return None
    projs = [float((c - curve_mid).DotProduct(w_unit)) for c in hc]
    w_min, w_max = min(projs), max(projs)
    half_d = 0.5 * max(bar_diam_ft, 1e-6)
    margin = cov_ft + half_d
    span = w_max - w_min
    if span < 2.0 * margin + 1e-6:
        return None
    w_lo = w_min + margin
    w_hi = w_max - margin
    if float(w_face_sign) > 0:
        w_bar = w_max - margin
    else:
        w_bar = w_min + margin
    if w_bar < w_lo - 1e-6 or w_bar > w_hi + 1e-6:
        return None
    return w_bar


def _curve_clr_base():
    """Tipo base CLR de Line para System.Array de curvas (CreateFromCurves)."""
    return clr.GetClrType(Line).BaseType


def _rebar_straight_from_line(document, host, bar_type, norm, ln):
    """Barra recta sin ganchos (hook = InvalidElementId). Prueba normales y orientaciones."""
    hid = ElementId.InvalidElementId
    ct = _curve_clr_base()
    arr = System.Array.CreateInstance(ct, 1)
    arr[0] = ln
    norms = [norm]
    try:
        norms.append(norm.Negate())
    except Exception:
        pass
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    for use_existing in (True, False):
        for create_new in (True, False):
            for nvec in norms:
                for so, eo in orient_pairs:
                    try:
                        r = Rebar.CreateFromCurves(
                            document,
                            RebarStyle.Standard,
                            bar_type,
                            hid,
                            hid,
                            host,
                            nvec,
                            arr,
                            so,
                            eo,
                            use_existing,
                            create_new,
                        )
                        if r:
                            return r
                    except Exception:
                        continue
    return None


def _rebar_straight_with_shape_fallback(
    document, host, bar_type, norm, ln, shapes, hook_type
):
    """
    Si CreateFromCurves (sin ganchos) falla en este proyecto/versión, intenta igual que las capas.
    Puede imponer forma/ganchos del RebarShape del proyecto.
    """
    if not shapes or not hook_type:
        return None
    curves = List[object]([ln])
    norms = [norm]
    try:
        norms.append(norm.Negate())
    except Exception:
        pass
    orient_pairs = _hook_orientation_pairs_create(False)
    for nvec in norms:
        for so, eo in orient_pairs:
            for shape in shapes:
                try:
                    r = Rebar.CreateFromCurvesAndShape(
                        document,
                        shape,
                        bar_type,
                        hook_type,
                        hook_type,
                        host,
                        nvec,
                        curves,
                        so,
                        eo,
                    )
                    if r:
                        return r
                except Exception:
                    continue
    return None


def _rebar_quantity(rebar):
    try:
        return int(rebar.Quantity)
    except Exception:
        try:
            return int(rebar.NumberOfBarPositions)
        except Exception:
            return 1


def _apply_fixed_number_layout(rebar, n_bars, array_length_ft):
    """
    Un solo elemento Rebar con layout fijo en cantidad (rebar set).
    Prueba varias combinaciones de barsOnNormalSide / include antes de FlipRebarSet,
    para evitar que la 2.ª barra quede fuera del host según el lado de propagación.
    """
    acc = rebar.GetShapeDrivenAccessor()
    if acc is None:
        return False
    try:
        if n_bars <= 1:
            acc.SetLayoutAsSingle()
            return True
    except Exception:
        return False

    def _qty_ok():
        try:
            return _rebar_quantity(rebar) == int(n_bars)
        except Exception:
            return False

    combos = (
        (True, True, True),
        (False, True, True),
        (True, False, False),
        (False, False, False),
    )

    def _try_all_combos():
        for b_side, inc0, inc1 in combos:
            try:
                acc.SetLayoutAsFixedNumber(
                    int(n_bars), float(array_length_ft), b_side, inc0, inc1
                )
                if _qty_ok():
                    return True
            except Exception:
                continue
        return False

    if _try_all_combos():
        return True
    try:
        acc.FlipRebarSet()
    except Exception:
        pass
    return _try_all_combos()


def _create_layer_rebar_set(
    document,
    host,
    p0,
    p1,
    axis,
    axis_off,
    width_dir,
    vert,
    norm_layout,
    n_section,
    bar_type,
    shapes,
    hook_type,
    es_capa_inferior,
    n_bars,
    w_ft,
    cov_ft,
    plan_shift,
    host_elems_for_hooks=None,
    axis_off_p0=None,
    axis_off_p1=None,
):
    """Un Rebar + SetLayoutAsFixedNumber; normal del plano = ancho → reparto a lo largo del ancho."""
    if n_bars < 1:
        return 0
    a0 = float(axis_off) if axis_off_p0 is None else float(axis_off_p0)
    a1 = float(axis_off) if axis_off_p1 is None else float(axis_off_p1)
    try:
        d_ft = float(bar_type.BarNominalDiameter)
        if d_ft <= 0:
            d_ft = float(getattr(bar_type, "BarModelDiameter", 0) or 0)
        if d_ft <= 0:
            d_ft = 0.04
    except Exception:
        d_ft = 0.04
    half = _width_half_span_ft(w_ft, cov_ft, d_ft)
    if n_bars <= 1:
        array_len = 0.0
        v_ref = 0.0
    else:
        array_len = 2.0 * half
        v_ref = -half
    q0 = p0 + axis * a0 + width_dir * v_ref + vert + plan_shift
    q1 = p1 - axis * a1 + width_dir * v_ref + vert + plan_shift
    try:
        # Sentido de la curva: Revit asocia ganchos inicio/fin al 1.er y 2.º extremo.
        # q0→q1 seguía el eje LocationCurve (p0→p1); en vigas suele voltear Left/Right
        # y dejar ambos ganchos hacia el exterior en sección. q1→q0 alinea ganchos hacia el interior.
        ln = Line.CreateBound(q1, q0)
    except Exception:
        return 0

    r = None
    norms = [norm_layout]
    try:
        norms.append(norm_layout.Negate())
    except Exception:
        pass
    orient_pairs = _hook_orientation_pairs_create(es_capa_inferior)
    curves = List[object]([ln])
    for nvec in norms:
        for so, eo in orient_pairs:
            for shape in shapes:
                try:
                    r = Rebar.CreateFromCurvesAndShape(
                        document,
                        shape,
                        bar_type,
                        hook_type,
                        hook_type,
                        host,
                        nvec,
                        curves,
                        so,
                        eo,
                    )
                    if r:
                        break
                except Exception:
                    continue
            if r:
                break
        if r:
            break

    if not r:
        return 0

    if not _apply_fixed_number_layout(r, n_bars, array_len):
        try:
            document.Delete(r.Id)
        except Exception:
            pass
        return 0

    _optimize_hook_orientations(
        r, host, n_section, es_capa_inferior, host_elems_for_hooks
    )

    return _rebar_quantity(r)


def _create_layer_rebar_set_try_hosts(
    document,
    hosts_try,
    p0,
    p1,
    axis,
    axis_off,
    width_dir,
    vert,
    norm_layout,
    n_section,
    bar_type,
    shapes,
    hook_type,
    es_capa_inferior,
    n_bars,
    w_ft,
    cov_ft,
    plan_shift,
    host_elems_for_hooks,
    axis_off_p0=None,
    axis_off_p1=None,
):
    """Misma creación probando varios posibles host (cadena de vigas)."""
    for host in hosts_try:
        if host is None:
            continue
        n = _create_layer_rebar_set(
            document,
            host,
            p0,
            p1,
            axis,
            axis_off,
            width_dir,
            vert,
            norm_layout,
            n_section,
            bar_type,
            shapes,
            hook_type,
            es_capa_inferior,
            n_bars,
            w_ft,
            cov_ft,
            plan_shift,
            host_elems_for_hooks,
            axis_off_p0,
            axis_off_p1,
        )
        if n > 0:
            return n
    return 0


def _create_lateral_rebar_set(
    document,
    host,
    p0,
    p1,
    axis,
    axis_off,
    w_lay,
    n_plane,
    w_face_sign,
    bar_type,
    n_bars,
    w_ft,
    d_ft,
    cov_ft,
    plan_shift,
    host_elems_for_hooks=None,
    axis_off_p0=None,
    axis_off_p1=None,
    shapes=None,
    hook_type=None,
    lateral_profile_elems=None,
):
    """
    Una instancia Rebar por cara lateral; reparto en canto con SetLayoutAsFixedNumber (normal = n).
    Ancho/canto desde bbox de vigas solamente (no columnas vecinas en la unión).
    Sin plan_shift. CreateFromCurves y fallback con forma del proyecto.
    """
    if n_bars < 1:
        return 0
    a0 = float(axis_off) if axis_off_p0 is None else float(axis_off_p0)
    a1 = float(axis_off) if axis_off_p1 is None else float(axis_off_p1)
    try:
        d_bar = float(bar_type.BarNominalDiameter)
        if d_bar <= 0:
            d_bar = float(getattr(bar_type, "BarModelDiameter", 0) or 0)
        if d_bar <= 0:
            d_bar = 0.04
    except Exception:
        d_bar = 0.04

    prof = (
        _dedupe_elements_by_id(lateral_profile_elems)
        if lateral_profile_elems is not None
        else _beam_elems_for_lateral_bbox(host_elems_for_hooks, host)
    )
    if not prof:
        prof = [host] if host else []

    _lat_clear_ft = _mm_to_ft(_LATERAL_CLEAR_FROM_FLEXURAL_MM)
    curve_mid_axis = p0 + axis * (0.5 * float((p1 - p0).DotProduct(axis)))
    bbox_w_pre = prof
    w_s_pre = _lateral_w_scalar_from_bbox(
        bbox_w_pre, curve_mid_axis, w_lay, w_face_sign, cov_ft, d_bar
    )
    if w_s_pre is None:
        w_s_pre = float(w_face_sign) * _width_half_span_ft(w_ft, cov_ft, d_bar)
    depth_pre = None
    if prof:
        depth_pre = _depth_span_along_n_from_bbox(
            prof,
            curve_mid_axis + w_lay * w_s_pre,
            n_plane,
            cov_ft,
            d_bar,
            _lat_clear_ft,
            _lat_clear_ft,
        )
    if depth_pre is None and host is not None:
        depth_pre = _depth_span_along_n_from_bbox(
            [host],
            curve_mid_axis + w_lay * w_s_pre,
            n_plane,
            cov_ft,
            d_bar,
            _lat_clear_ft,
            _lat_clear_ft,
        )
    n_mid_pre = depth_pre[0] if depth_pre else 0.0
    lateral_off_trim = w_lay * w_s_pre + n_plane * n_mid_pre

    pa, pb = _axis_cover_trim_endpoints(
        p0, p1, host, cov_ft, d_bar, lateral_off_trim
    )
    if pa is not None and pb is not None:
        p_axis0, p_axis1 = pa, pb
    else:
        m_end = cov_ft + 0.5 * max(d_bar, 1e-6)
        p_axis0 = p0 + axis * max(a0, m_end)
        p_axis1 = p1 - axis * max(a1, m_end)

    curve_mid = p_axis0 + (p_axis1 - p_axis0) * 0.5
    depth_res = None
    if prof:
        depth_res = _depth_span_along_n_from_bbox(
            prof,
            curve_mid,
            n_plane,
            cov_ft,
            d_bar,
            _lat_clear_ft,
            _lat_clear_ft,
        )
    if depth_res is None and host is not None:
        depth_res = _depth_span_along_n_from_bbox(
            [host],
            curve_mid,
            n_plane,
            cov_ft,
            d_bar,
            _lat_clear_ft,
            _lat_clear_ft,
        )
    if depth_res is None:
        n_strip_mid = 0.0
        half_raw = _depth_half_span_ft(d_ft, cov_ft, d_bar)
        inner_lat = max(0.0, 2.0 * half_raw - 2.0 * _lat_clear_ft)
        usable_half = 0.5 * inner_lat
    else:
        n_strip_mid, usable_half = depth_res

    bbox_w = prof
    w_s = _lateral_w_scalar_from_bbox(
        bbox_w, curve_mid, w_lay, w_face_sign, cov_ft, d_bar
    )
    if w_s is None:
        w_s = float(w_face_sign) * _width_half_span_ft(w_ft, cov_ft, d_bar)

    # plan_shift (centro bbox en ancho) empuja la línea fuera de la cara lateral; no usar aquí.
    _lat_shift = XYZ(0, 0, 0)

    if n_bars <= 1:
        array_len = 0.0
        n_ref = 0.0
    else:
        array_len = 2.0 * usable_half
        n_ref = -usable_half

    n_off = n_strip_mid + n_ref
    q0 = p_axis0 + w_lay * w_s + n_plane * n_off + _lat_shift
    q1 = p_axis1 + w_lay * w_s + n_plane * n_off + _lat_shift
    try:
        ln = Line.CreateBound(q1, q0)
    except Exception:
        return 0

    shapes_lat = _rebar_shapes_simple_line_only(shapes)
    if not shapes_lat:
        shapes_lat = shapes

    r = _rebar_straight_from_line(document, host, bar_type, n_plane, ln)
    if not r:
        r = _rebar_straight_with_shape_fallback(
            document, host, bar_type, n_plane, ln, shapes_lat, hook_type
        )
    if not r:
        return 0

    if n_bars > 1:
        acc0 = None
        try:
            acc0 = r.GetShapeDrivenAccessor()
        except Exception:
            pass
        if acc0 is None and shapes_lat and hook_type:
            try:
                document.Delete(r.Id)
            except Exception:
                pass
            r = _rebar_straight_with_shape_fallback(
                document, host, bar_type, n_plane, ln, shapes_lat, hook_type
            )
            if not r:
                return 0

    if not _apply_fixed_number_layout(r, n_bars, array_len):
        try:
            document.Delete(r.Id)
        except Exception:
            pass
        return 0

    try:
        wh = w_lay.Multiply(float(w_face_sign))
    except Exception:
        wh = w_lay * float(w_face_sign)
    _optimize_hook_orientations(r, host, wh, False, prof)

    return _rebar_quantity(r)


def _create_lateral_rebar_set_try_hosts(
    document,
    hosts_try,
    p0,
    p1,
    axis,
    axis_off,
    w_lay,
    n_plane,
    w_face_sign,
    bar_type,
    n_bars,
    w_ft,
    d_ft,
    cov_ft,
    plan_shift,
    host_elems_for_hooks,
    axis_off_p0=None,
    axis_off_p1=None,
    shapes=None,
    hook_type=None,
    lateral_profile_elems=None,
):
    for host in hosts_try:
        if host is None:
            continue
        n = _create_lateral_rebar_set(
            document,
            host,
            p0,
            p1,
            axis,
            axis_off,
            w_lay,
            n_plane,
            w_face_sign,
            bar_type,
            n_bars,
            w_ft,
            d_ft,
            cov_ft,
            plan_shift,
            host_elems_for_hooks,
            axis_off_p0,
            axis_off_p1,
            shapes,
            hook_type,
            lateral_profile_elems,
        )
        if n > 0:
            return n
    return 0


def _place_layers_on_beam(
    document,
    elem,
    bar_type_sup,
    bar_type_inf,
    n_sup,
    n_inf,
    shapes,
    hook_type,
):
    cover_mm = _COVER_MM_FIXED
    loc = elem.Location
    if not isinstance(loc, LocationCurve):
        return 0, u"Sin LocationCurve."
    curve = loc.Curve
    frame = _beam_frame(curve)
    if not frame:
        return 0, u"Viga casi vertical o curva inválida."
    axis, width_dir, _, p0, p1, beam_len = frame
    if beam_len <= 2.0 * _mm_to_ft(cover_mm) + 1e-4:
        return 0, u"Viga demasiado corta para el recubrimiento."

    w_ft, d_ft = _read_width_depth_ft(document, elem, curve)
    cov = _mm_to_ft(cover_mm)
    axis_unit = axis.Normalize()
    p0, p1, neigh, ext0_ft, ext1_ft = _extend_span_to_adjacent_supports(
        document, elem, p0, p1, axis_unit, cov, None
    )
    _ext_eps = 1e-5
    axis_off_p0 = 0.0 if ext0_ft > _ext_eps else cov
    axis_off_p1 = 0.0 if ext1_ft > _ext_eps else cov
    try:
        line_adj = Line.CreateBound(p0, p1)
    except Exception:
        return 0, u"Eje inválido tras extender apoyos."
    frame2 = _beam_frame(line_adj)
    if not frame2:
        return 0, u"Viga casi vertical o curva inválida."
    axis, width_dir, _, p0, p1, beam_len = frame2
    if beam_len <= 2.0 * cov + 1e-4:
        return 0, u"Viga demasiado corta para el recubrimiento."
    axis_off = cov
    if beam_len <= 2.0 * axis_off + 1e-4:
        return 0, u"Longitud útil insuficiente."

    ortho = _rebar_ortho_frame_t_w_n(axis, width_dir)
    if ortho is None:
        return 0, u"No se pudo calcular triedro eje/ancho/canto."
    t_axis, w_lay, n_plane = ortho

    # Unión viga+columnas solo para host/ganchos; alturas y planta solo desde la(s) viga(s).
    # Si dz usara bbox de columna, z_inf podría quedar bajo el fondo real de la viga y Revit
    # rechaza la capa inferior con host=viga.
    bbox_elems = _dedupe_elements_by_id([elem] + neigh)

    # Vigas casi horizontales: alturas de capa desde BoundingBox real (caras inf/sup del sólido),
    # no desde d/2 respecto a la LocationCurve (suele ser ref. distinta al centro de sección).
    use_bbox_z = False
    dz_inf = dz_sup = None
    if abs(t_axis.Z) < 0.4:
        dz_pair = _layer_dz_world_z_from_bbox(elem, p0, p1, cov)
        if dz_pair is not None:
            dz_inf, dz_sup = dz_pair
            use_bbox_z = True

    # Normal a CreateFromCurvesAndShape = ancho: Revit extiende el set a lo largo de ±normal.
    if w_lay is None or w_lay.GetLength() < 1e-9:
        return 0, u"No se pudo calcular la normal de armadura (ancho)."
    norm_layout = w_lay

    plan_shift = _shift_line_mid_along_width_to_bbox_center(elem, p0, p1, w_lay, w_ft)

    hosts_try = _dedupe_elements_by_id(neigh + [elem])

    errs = []

    def run_layer(n_bars, bar_type, dz_world_z, sign_depth, es_capa_inferior):
        if n_bars < 1 or bar_type is None:
            return 0
        if dz_world_z is not None:
            vert = XYZ(0, 0, float(dz_world_z))
        else:
            z_along_depth = sign_depth * (0.5 * d_ft - cov)
            vert = n_plane * z_along_depth
        n = _create_layer_rebar_set_try_hosts(
            document,
            hosts_try,
            p0,
            p1,
            t_axis,
            axis_off,
            w_lay,
            vert,
            norm_layout,
            n_plane,
            bar_type,
            shapes,
            hook_type,
            es_capa_inferior,
            n_bars,
            w_ft,
            cov,
            plan_shift,
            bbox_elems,
            axis_off_p0,
            axis_off_p1,
        )
        if n <= 0:
            errs.append(u"No CreateFromCurvesAndShape o layout fijo.")
        return n

    creados = (
        run_layer(n_inf, bar_type_inf, dz_inf if use_bbox_z else None, -1.0, True)
        + run_layer(n_sup, bar_type_sup, dz_sup if use_bbox_z else None, 1.0, False)
    )

    msg = None
    if errs and creados == 0:
        msg = errs[0][:120] if errs else u"Error desconocido."
    return creados, msg


def _build_collinear_chains_from_elements(document, elems):
    """Lista de cadenas; cada cadena es vigas colineales ordenadas a lo largo del eje."""
    bd = []
    for e in elems:
        d = _beam_placement_data(document, e)
        if d:
            bd.append(d)
    if not bd:
        return []
    chains_idx = _cluster_beam_indices(bd)
    out = []
    for g in chains_idx:
        ref = bd[g[0]]
        ordered = _order_cluster_indices(bd, g, ref["axis"], ref["p0"])
        out.append([bd[i]["elem"] for i in ordered])
    return out


def _place_layers_on_beam_aligned_chain(
    document,
    chain_elems,
    bar_type_sup,
    bar_type_inf,
    n_sup,
    n_inf,
    shapes,
    hook_type,
):
    """
    Un Rebar por capa en toda la cadena (curva de un extremo al otro).
    Si ningún host admite la curva completa, coloca por viga.
    """
    cover_mm = _COVER_MM_FIXED
    if not chain_elems:
        return 0, u"Cadena vacía."
    if len(chain_elems) == 1:
        return _place_layers_on_beam(
            document,
            chain_elems[0],
            bar_type_sup,
            bar_type_inf,
            n_sup,
            n_inf,
            shapes,
            hook_type,
        )

    bd = [_beam_placement_data(document, e) for e in chain_elems]
    if any(x is None for x in bd):
        return 0, u"Cadena: viga sin LocationCurve o eje inválido."

    axis0 = bd[0]["axis"]
    for d in bd[1:]:
        if abs(float(d["axis"].DotProduct(axis0))) < 0.97:
            return 0, u"Cadena: vigas no paralelas."

    ref_p0 = bd[0]["p0"]
    axis_use = axis0.Normalize()
    all_pts = []
    for d in bd:
        all_pts.extend([d["p0"], d["p1"]])

    def _proj_along(p):
        return float((p - ref_p0).DotProduct(axis_use))

    p_lo = min(all_pts, key=_proj_along)
    p_hi = max(all_pts, key=_proj_along)

    cov = _mm_to_ft(cover_mm)
    axis_seg = p_hi - p_lo
    if axis_seg.GetLength() < 1e-9:
        return 0, u"Cadena: longitud nula."
    axis_unit = axis_seg.Normalize()
    excl = set()
    for e in chain_elems:
        try:
            excl.add(int(e.Id.IntegerValue))
        except Exception:
            pass
    p_lo, p_hi, neigh, ext_lo_ft, ext_hi_ft = _extend_span_to_adjacent_supports(
        document, chain_elems[0], p_lo, p_hi, axis_unit, cov, excl
    )
    _ext_eps = 1e-5
    axis_off_p0 = 0.0 if ext_lo_ft > _ext_eps else cov
    axis_off_p1 = 0.0 if ext_hi_ft > _ext_eps else cov

    try:
        line_merged = Line.CreateBound(p_lo, p_hi)
    except Exception:
        return 0, u"Cadena: no se pudo unir extremos."

    frame_merged = _beam_frame(line_merged)
    if not frame_merged:
        return 0, u"Cadena: eje casi vertical."

    axis_m, width_dir_m, depth_dir_m, p0m, p1m, beam_len = frame_merged
    axis_off = cov
    if beam_len <= 2.0 * cov + 1e-4:
        return 0, u"Cadena demasiado corta para el recubrimiento."
    if beam_len <= 2.0 * axis_off + 1e-4:
        return 0, u"Cadena: longitud útil insuficiente."

    w_ft = 0.0
    d_ft = 0.0
    for e in chain_elems:
        try:
            c = e.Location.Curve
            wi, di = _read_width_depth_ft(document, e, c)
            w_ft = max(w_ft, wi)
            d_ft = max(d_ft, di)
        except Exception:
            continue

    ortho = _rebar_ortho_frame_t_w_n(axis_m, width_dir_m)
    if ortho is None:
        return 0, u"Cadena: no se pudo calcular triedro."
    t_axis, w_lay, n_plane = ortho
    norm_layout = w_lay

    # dz y planta: solo vigas de la cadena (no columnas vecinas): evita fibra inf en el fondo
    # del bbox de columna y fallo de la capa inferior con host=viga.
    bbox_elems = _dedupe_elements_by_id(chain_elems + neigh)

    use_bbox_z = False
    dz_inf = dz_sup = None
    if abs(t_axis.Z) < 0.4:
        dz_pair = _layer_dz_world_z_from_bbox_multi(chain_elems, p0m, p1m, cov)
        if dz_pair is not None:
            dz_inf, dz_sup = dz_pair
            use_bbox_z = True

    plan_shift = _shift_line_mid_along_width_to_bbox_union(
        chain_elems, p0m, p1m, w_lay, w_ft
    )

    hosts_try = _dedupe_elements_by_id(
        neigh + sorted(chain_elems, key=_beam_curve_length_ft, reverse=True)
    )

    errs = []

    def run_layer(n_bars, bar_type, dz_world_z, sign_depth, es_capa_inferior):
        if n_bars < 1 or bar_type is None:
            return 0
        if dz_world_z is not None:
            vert = XYZ(0, 0, float(dz_world_z))
        else:
            z_along_depth = sign_depth * (0.5 * d_ft - cov)
            vert = n_plane * z_along_depth
        n = _create_layer_rebar_set_try_hosts(
            document,
            hosts_try,
            p0m,
            p1m,
            t_axis,
            axis_off,
            w_lay,
            vert,
            norm_layout,
            n_plane,
            bar_type,
            shapes,
            hook_type,
            es_capa_inferior,
            n_bars,
            w_ft,
            cov,
            plan_shift,
            bbox_elems,
            axis_off_p0,
            axis_off_p1,
        )
        if n <= 0:
            errs.append(u"Create/layout en cadena.")
        return n

    creados = (
        run_layer(n_inf, bar_type_inf, dz_inf if use_bbox_z else None, -1.0, True)
        + run_layer(n_sup, bar_type_sup, dz_sup if use_bbox_z else None, 1.0, False)
    )

    if creados > 0:
        return creados, None

    sub = 0
    for e in chain_elems:
        ne, _ = _place_layers_on_beam(
            document,
            e,
            bar_type_sup,
            bar_type_inf,
            n_sup,
            n_inf,
            shapes,
            hook_type,
        )
        sub += ne
    if sub > 0:
        return sub, u"Cadena no admitida como un solo Rebar; colocado por tramo."
    return 0, (errs[0][:120] if errs else u"Error en cadena.")


def _place_lateral_on_beam(
    document,
    elem,
    bar_type_lat,
    n_lat,
    shapes=None,
    hook_type=None,
):
    """
    Barras laterales en ambas caras del alma (±w); reparto en altura.
    Eje solo en la viga; posición en ancho/canto desde bbox; fallback con forma del proyecto.
    """
    if n_lat < 1 or bar_type_lat is None:
        return 0, None
    cover_mm = _COVER_MM_FIXED
    loc = elem.Location
    if not isinstance(loc, LocationCurve):
        return 0, u"Sin LocationCurve."
    curve = loc.Curve
    frame = _beam_frame(curve)
    if not frame:
        return 0, u"Viga casi vertical o curva inválida."
    axis, width_dir, _, p0, p1, beam_len = frame
    if beam_len <= 2.0 * _mm_to_ft(cover_mm) + 1e-4:
        return 0, u"Viga demasiado corta para el recubrimiento."

    w_ft, d_ft = _read_width_depth_ft(document, elem, curve)
    cov = _mm_to_ft(cover_mm)
    axis_unit = axis.Normalize()
    p0, p1, neigh, ext0_ft, ext1_ft = _extend_span_to_adjacent_supports(
        document, elem, p0, p1, axis_unit, cov, None
    )
    _ext_eps = 1e-5
    axis_off_p0 = 0.0 if ext0_ft > _ext_eps else cov
    axis_off_p1 = 0.0 if ext1_ft > _ext_eps else cov
    try:
        line_adj = Line.CreateBound(p0, p1)
    except Exception:
        return 0, u"Eje inválido tras extender apoyos."
    frame2 = _beam_frame(line_adj)
    if not frame2:
        return 0, u"Viga casi vertical o curva inválida."
    axis, width_dir, _, p0, p1, beam_len = frame2
    if beam_len <= 2.0 * cov + 1e-4:
        return 0, u"Viga demasiado corta para el recubrimiento."
    axis_off = cov
    if beam_len <= 2.0 * axis_off + 1e-4:
        return 0, u"Longitud útil insuficiente."

    ortho = _rebar_ortho_frame_t_w_n(axis, width_dir)
    if ortho is None:
        return 0, u"No se pudo calcular triedro eje/ancho/canto."
    t_axis, w_lay, n_plane = ortho

    bbox_elems = _dedupe_elements_by_id([elem] + neigh)
    hosts_try = _dedupe_elements_by_id(neigh + [elem])
    lateral_plan = XYZ(0, 0, 0)
    lateral_prof = _dedupe_elements_by_id([elem])

    creados = 0
    errs = []
    for sgn in (1.0, -1.0):
        n = _create_lateral_rebar_set_try_hosts(
            document,
            hosts_try,
            p0,
            p1,
            t_axis,
            axis_off,
            w_lay,
            n_plane,
            sgn,
            bar_type_lat,
            n_lat,
            w_ft,
            d_ft,
            cov,
            lateral_plan,
            bbox_elems,
            axis_off_p0,
            axis_off_p1,
            shapes,
            hook_type,
            lateral_prof,
        )
        if n > 0:
            creados += n
        else:
            errs.append(u"Cara lateral sign={}".format(int(sgn)))

    msg = None
    if errs and creados == 0:
        msg = u"No se pudieron crear laterales ({}).".format(u", ".join(errs))
    elif errs and creados > 0:
        msg = u"Parcial: {}".format(u", ".join(errs))
    return creados, msg


def _place_lateral_on_beam_aligned_chain(
    document,
    chain_elems,
    bar_type_lat,
    n_lat,
    shapes=None,
    hook_type=None,
):
    """
    Vigas colineales: misma geometría que capas sup/inf (_place_layers_on_beam_aligned_chain):
    eje fusionado de extremo a extremo, prolongación en apoyos, hosts_try y bbox de cadena+vecinos.
    Dos Rebar (uno por cara lateral ±w), no dos por cada viga del tramo.
    Si ningún host admite la curva completa, coloca por viga.
    """
    if n_lat < 1 or bar_type_lat is None:
        return 0, None
    if not chain_elems:
        return 0, u"Cadena vacía."
    if len(chain_elems) == 1:
        return _place_lateral_on_beam(
            document,
            chain_elems[0],
            bar_type_lat,
            n_lat,
            shapes,
            hook_type,
        )

    bd = [_beam_placement_data(document, e) for e in chain_elems]
    if any(x is None for x in bd):
        return 0, u"Cadena: viga sin LocationCurve o eje inválido."

    axis0 = bd[0]["axis"]
    for d in bd[1:]:
        if abs(float(d["axis"].DotProduct(axis0))) < 0.97:
            return 0, u"Cadena: vigas no paralelas."

    ref_p0 = bd[0]["p0"]
    axis_use = axis0.Normalize()
    all_pts = []
    for d in bd:
        all_pts.extend([d["p0"], d["p1"]])

    def _proj_along(p):
        return float((p - ref_p0).DotProduct(axis_use))

    p_lo = min(all_pts, key=_proj_along)
    p_hi = max(all_pts, key=_proj_along)

    cov = _mm_to_ft(_COVER_MM_FIXED)
    axis_seg = p_hi - p_lo
    if axis_seg.GetLength() < 1e-9:
        return 0, u"Cadena: longitud nula."
    axis_unit = axis_seg.Normalize()
    excl = set()
    for e in chain_elems:
        try:
            excl.add(int(e.Id.IntegerValue))
        except Exception:
            pass
    p_lo, p_hi, neigh, ext_lo_ft, ext_hi_ft = _extend_span_to_adjacent_supports(
        document, chain_elems[0], p_lo, p_hi, axis_unit, cov, excl
    )
    _ext_eps = 1e-5
    axis_off_p0 = 0.0 if ext_lo_ft > _ext_eps else cov
    axis_off_p1 = 0.0 if ext_hi_ft > _ext_eps else cov

    try:
        line_merged = Line.CreateBound(p_lo, p_hi)
    except Exception:
        return 0, u"Cadena: no se pudo unir extremos."

    frame_merged = _beam_frame(line_merged)
    if not frame_merged:
        return 0, u"Cadena: eje casi vertical."

    axis_m, width_dir_m, depth_dir_m, p0m, p1m, beam_len = frame_merged
    axis_off = cov
    if beam_len <= 2.0 * cov + 1e-4:
        return 0, u"Cadena demasiado corta para el recubrimiento."
    if beam_len <= 2.0 * axis_off + 1e-4:
        return 0, u"Cadena: longitud útil insuficiente."

    w_ft = 0.0
    d_ft = 0.0
    for e in chain_elems:
        try:
            c = e.Location.Curve
            wi, di = _read_width_depth_ft(document, e, c)
            w_ft = max(w_ft, wi)
            d_ft = max(d_ft, di)
        except Exception:
            continue

    ortho = _rebar_ortho_frame_t_w_n(axis_m, width_dir_m)
    if ortho is None:
        return 0, u"Cadena: no se pudo calcular triedro."
    t_axis, w_lay, n_plane = ortho

    bbox_elems = _dedupe_elements_by_id(chain_elems + neigh)
    lateral_plan = XYZ(0, 0, 0)
    lateral_prof = _dedupe_elements_by_id(chain_elems)

    hosts_try = _dedupe_elements_by_id(
        neigh + sorted(chain_elems, key=_beam_curve_length_ft, reverse=True)
    )

    errs = []
    creados = 0
    for sgn in (1.0, -1.0):
        n = _create_lateral_rebar_set_try_hosts(
            document,
            hosts_try,
            p0m,
            p1m,
            t_axis,
            axis_off,
            w_lay,
            n_plane,
            sgn,
            bar_type_lat,
            n_lat,
            w_ft,
            d_ft,
            cov,
            lateral_plan,
            bbox_elems,
            axis_off_p0,
            axis_off_p1,
            shapes,
            hook_type,
            lateral_prof,
        )
        if n > 0:
            creados += n
        else:
            errs.append(u"Cara lateral sign={}".format(int(sgn)))

    if creados > 0:
        msg = None
        if errs:
            msg = u"Parcial: {}".format(u", ".join(errs))
        return creados, msg

    sub = 0
    fer = []
    for e in chain_elems:
        ne, el = _place_lateral_on_beam(
            document, e, bar_type_lat, n_lat, shapes, hook_type
        )
        sub += ne
        if ne == 0 and el:
            try:
                fer.append(u"{} {}".format(int(e.Id.IntegerValue), el))
            except Exception:
                fer.append(el or u"?")
    if sub > 0:
        return (
            sub,
            u"Cadena lateral no admitida como un solo Rebar por cara; colocado por tramo.",
        )
    return 0, (
        (u"; ".join(fer)[:320] if fer else None)
        or (u", ".join(errs) if errs else u"Error en cadena.")
    )


def _parse_positive_int(s, default, label):
    try:
        t = (s or u"").strip()
        if not t:
            return default
        n = int(t)
        if n < 1:
            raise ValueError()
        return n
    except Exception:
        raise ValueError(u"{}: entero ≥ 1.".format(label))


class SeleccionarVigasHandler(IExternalEventHandler):
    """Selección de vigas en contexto API (mismo patrón que Area Reinforcement Muro RPS)."""

    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        uidoc = uiapp.ActiveUIDocument
        doc = uidoc.Document
        win = self._window_ref()
        if not win:
            return
        try:
            win._document = doc
            refs = list(
                uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    VigasArmaduraFilter(),
                    u"Selecciona vigas (Structural Framing — Beam). Finaliza con Finish o Cancel.",
                )
            )
            if not refs:
                win._beam_ids = []
                win._set_estado(u"Selección vacía.")
                win._set_viga_info(u"Ninguna viga seleccionada.")
                return

            seen = set()
            beam_ids = []
            for ref in refs:
                elem = doc.GetElement(ref.ElementId)
                if elem is None:
                    continue
                try:
                    eid = int(elem.Id.IntegerValue)
                except Exception:
                    continue
                if eid in seen:
                    continue
                seen.add(eid)
                beam_ids.append(elem.Id)

            if not beam_ids:
                win._beam_ids = []
                win._set_viga_info(u"Ningún elemento válido (solo vigas).")
                TaskDialog.Show(
                    u"Armadura en vigas",
                    u"Los elementos seleccionados no son vigas estructurales (Beam).",
                )
                return

            win._beam_ids = beam_ids
            n = len(beam_ids)
            if n == 1:
                el = doc.GetElement(beam_ids[0])
                try:
                    tid = el.Id.IntegerValue
                except Exception:
                    tid = u"?"
                win._set_viga_info(u"1 viga seleccionada (ID: {}).".format(tid))
            else:
                ids_txt = u", ".join(str(beam_ids[i].IntegerValue) for i in range(min(n, 5)))
                if n > 5:
                    ids_txt += u"..."
                win._set_viga_info(u"{} vigas seleccionadas (IDs: {}).".format(n, ids_txt))
            win._set_estado(u"Listo para colocar armaduras.")
            try:
                win._refresh_txt_n_lat_from_selection()
            except Exception:
                pass
        except Exception as ex:
            err = str(ex).lower()
            if "cancel" not in err and "operation" not in err:
                win._set_estado(u"Error: {}.".format(str(ex)))
                TaskDialog.Show(u"Armadura en vigas — Error", str(ex))
            else:
                win._set_estado(u"Selección cancelada.")
        finally:
            try:
                win._win.Show()
                win._win.Activate()
            except Exception:
                pass

    def GetName(self):
        return u"SeleccionarVigasArmaduraCapas"


class ColocarArmaduraVigasHandler(IExternalEventHandler):
    """Crea barras en transacción (contexto API)."""

    def __init__(self, window_ref):
        self._window_ref = window_ref
        self.opts = None
        self.beam_ids = None

    def Execute(self, uiapp):
        doc = uiapp.ActiveUIDocument.Document
        win = self._window_ref()
        if not win:
            return
        opts = self.opts
        beam_ids = self.beam_ids
        self.opts = None
        self.beam_ids = None
        if not opts or not beam_ids:
            return

        bar_sup = doc.GetElement(opts["id_sup"])
        bar_inf = doc.GetElement(opts["id_inf"])
        if not isinstance(bar_sup, RebarBarType) or not isinstance(bar_inf, RebarBarType):
            TaskDialog.Show(u"Armadura en vigas", u"Tipo de barra inválido.")
            return

        bar_lat = None
        if opts.get("laterales"):
            bar_lat = doc.GetElement(opts["id_lat"])
            if not isinstance(bar_lat, RebarBarType):
                TaskDialog.Show(
                    u"Armadura en vigas",
                    u"Tipo de barra inválido para barras laterales.",
                )
                return

        shapes = _pick_rebar_shapes(doc)
        if not shapes:
            TaskDialog.Show(
                u"Armadura en vigas",
                u"No hay RebarShape estándar en el proyecto.",
            )
            return

        hook_type = _pick_hook_90(doc)
        if not hook_type:
            TaskDialog.Show(
                u"Armadura en vigas",
                u"No se encontró un RebarHookType de 90°. Crea uno en el proyecto.",
            )
            return

        elems_sel = []
        for eid in beam_ids:
            elem = doc.GetElement(eid)
            if elem is not None:
                elems_sel.append(elem)

        chains = _build_collinear_chains_from_elements(doc, elems_sel)
        if not chains:
            TaskDialog.Show(
                u"Armadura en vigas",
                u"No hay vigas con eje válido en la selección.",
            )
            return

        total = 0
        detalles = []
        t = Transaction(
            doc,
            u"BIMTools — Armadura en vigas (capas y laterales)",
        )
        t.Start()
        try:
            for chain in chains:
                if len(chain) == 1:
                    elem = chain[0]
                    n, err = _place_layers_on_beam(
                        doc,
                        elem,
                        bar_sup,
                        bar_inf,
                        opts["n_sup"],
                        opts["n_inf"],
                        shapes,
                        hook_type,
                    )
                else:
                    n, err = _place_layers_on_beam_aligned_chain(
                        doc,
                        chain,
                        bar_sup,
                        bar_inf,
                        opts["n_sup"],
                        opts["n_inf"],
                        shapes,
                        hook_type,
                    )
                total += n
                if err:
                    if n == 0:
                        try:
                            ids = u", ".join(
                                unicode(e.Id.IntegerValue) for e in chain[:6]
                            )
                        except Exception:
                            ids = u"?"
                        detalles.append(u"[{}]: {}".format(ids, err))
                    else:
                        detalles.append(err)

                if opts.get("laterales") and bar_lat is not None:
                    if len(chain) == 1:
                        nl, el = _place_lateral_on_beam(
                            doc,
                            chain[0],
                            bar_lat,
                            opts["n_lat"],
                            shapes,
                            hook_type,
                        )
                    else:
                        nl, el = _place_lateral_on_beam_aligned_chain(
                            doc,
                            chain,
                            bar_lat,
                            opts["n_lat"],
                            shapes,
                            hook_type,
                        )
                    total += nl
                    if el:
                        if nl == 0:
                            try:
                                ids = u", ".join(
                                    unicode(e.Id.IntegerValue) for e in chain[:6]
                                )
                            except Exception:
                                ids = u"?"
                            detalles.append(u"[{}] laterales: {}".format(ids, el))
                        else:
                            detalles.append(u"Laterales: {}".format(el))
            t.Commit()
        except Exception as ex:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            try:
                win._set_estado(u"Error al crear barras.")
            except Exception:
                pass
            TaskDialog.Show(u"Armadura en vigas — error", unicode(ex))
            try:
                win._win.Activate()
            except Exception:
                pass
            return

        msg = u"Barras creadas: {}.".format(total)
        if detalles:
            msg += u"\n\nSin crear en:\n" + u"\n".join(detalles[:8])
            if len(detalles) > 8:
                msg += u"\n..."
        try:
            win._set_estado(u"Última operación: {} barras.".format(total))
        except Exception:
            pass
        TaskDialog.Show(u"Armadura en vigas", msg)
        try:
            win._win.Activate()
        except Exception:
            pass

    def GetName(self):
        return u"ColocarArmaduraVigasCapas"


class ArmaduraVigasCapasWindow(object):
    """Ventana no modal: configuración, selección de vigas (ExternalEvent) y colocación (ExternalEvent)."""

    def __init__(self, revit, manual_dir, bar_type_rows):
        self._revit = revit
        self._manual_dir = manual_dir
        self._document = None
        self._beam_ids = []
        self._win = XamlReader.Parse(XAML)

        self._seleccion_handler = SeleccionarVigasHandler(weakref.ref(self))
        self._seleccion_event = ExternalEvent.Create(self._seleccion_handler)
        self._colocar_handler = ColocarArmaduraVigasHandler(weakref.ref(self))
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)

        cmb_s = self._win.FindName("CmbSup")
        cmb_i = self._win.FindName("CmbInf")
        cmb_l = self._win.FindName("CmbLat")
        for label, eid in bar_type_rows:
            it_s = ComboBoxItem()
            it_s.Content = label
            it_s.Tag = eid
            cmb_s.Items.Add(it_s)
            it_i = ComboBoxItem()
            it_i.Content = label
            it_i.Tag = eid
            cmb_i.Items.Add(it_i)
            if cmb_l is not None:
                it_l = ComboBoxItem()
                it_l.Content = label
                it_l.Tag = eid
                cmb_l.Items.Add(it_l)
        if cmb_s.Items.Count:
            cmb_s.SelectedIndex = 0
        if cmb_i.Items.Count:
            cmb_i.SelectedIndex = min(1, cmb_i.Items.Count - 1) if cmb_i.Items.Count > 1 else 0
        if cmb_l is not None and cmb_l.Items.Count:
            cmb_l.SelectedIndex = 0

        self._win.FindName("BtnSeleccionar").Click += RoutedEventHandler(self._on_seleccionar)
        self._win.FindName("BtnColocar").Click += RoutedEventHandler(self._on_colocar)
        self._win.FindName("BtnCancel").Click += RoutedEventHandler(self._on_cerrar)
        self._win.FindName("BtnManual").Click += RoutedEventHandler(self._on_manual)

        def _on_close_cmd(sender, e):
            try:
                self._win.Close()
            except Exception:
                pass

        self._win.CommandBindings.Add(CommandBinding(ApplicationCommands.Close, _on_close_cmd))
        self._win.InputBindings.Add(KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None))

    def _set_estado(self, msg):
        try:
            txt = self._win.FindName("TxtEstado")
            if txt:
                txt.Text = msg or u""
        except Exception:
            pass

    def _set_viga_info(self, msg):
        try:
            txt = self._win.FindName("TxtVigaInfo")
            if txt:
                txt.Text = msg or u""
        except Exception:
            pass

    def _suggested_n_lat(self):
        if not self._document or not self._beam_ids:
            return 1
        try:
            return _default_n_lat_from_beams(self._document, self._beam_ids)
        except Exception:
            return 1

    def _refresh_txt_n_lat_from_selection(self):
        n = self._suggested_n_lat()
        try:
            tb = self._win.FindName("TxtNlat")
            if tb is not None:
                tb.Text = unicode(n)
        except Exception:
            pass

    def _parse_options_from_ui(self):
        cmb_s = self._win.FindName("CmbSup")
        cmb_i = self._win.FindName("CmbInf")
        cmb_l = self._win.FindName("CmbLat")
        chk_lat = self._win.FindName("ChkLaterales")
        ns = _parse_positive_int(self._win.FindName("TxtNsup").Text, 2, u"Capa superior")
        ni = _parse_positive_int(self._win.FindName("TxtNinf").Text, 2, u"Capa inferior")
        si = cmb_s.SelectedItem if cmb_s else None
        ii = cmb_i.SelectedItem if cmb_i else None
        if si is None or ii is None:
            raise ValueError(u"Selecciona tipo de barra en ambos combos.")
        out = {
            "n_sup": ns,
            "n_inf": ni,
            "id_sup": si.Tag,
            "id_inf": ii.Tag,
            "laterales": False,
            "n_lat": 0,
            "id_lat": None,
        }
        if chk_lat is not None and chk_lat.IsChecked:
            nlat = _parse_positive_int(
                self._win.FindName("TxtNlat").Text,
                self._suggested_n_lat(),
                u"Laterales (nº barras)",
            )
            li = cmb_l.SelectedItem if cmb_l else None
            if li is None:
                raise ValueError(u"Selecciona tipo de barra para laterales.")
            out["laterales"] = True
            out["n_lat"] = nlat
            out["id_lat"] = li.Tag
        return out

    def _on_seleccionar(self, sender, args):
        self._win.Hide()
        self._seleccion_event.Raise()

    def _on_colocar(self, sender, args):
        try:
            opts = self._parse_options_from_ui()
        except ValueError as ex:
            MessageBox.Show(
                unicode(ex),
                u"Armadura en vigas",
                MessageBoxButton.OK,
                MessageBoxImage.Warning,
            )
            return
        if not self._beam_ids:
            MessageBox.Show(
                u"Primero selecciona una o más vigas con «Seleccionar vigas en modelo».",
                u"Armadura en vigas",
                MessageBoxButton.OK,
                MessageBoxImage.Warning,
            )
            self._set_estado(u"Selecciona vigas antes de colocar armaduras.")
            return
        self._colocar_handler.opts = opts
        self._colocar_handler.beam_ids = list(self._beam_ids)
        self._set_estado(u"Colocando armaduras…")
        self._colocar_event.Raise()

    def _on_cerrar(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass

    def _on_manual(self, sender, args):
        if not self._manual_dir:
            MessageBox.Show(
                u"Ruta del manual no configurada.",
                u"Armadura en vigas",
                MessageBoxButton.OK,
                MessageBoxImage.Information,
            )
            return
        manual_path = os.path.join(self._manual_dir, "manual_usuario.html")
        try:
            if os.path.isfile(manual_path):
                os.startfile(os.path.abspath(manual_path))
            else:
                MessageBox.Show(
                    u"No se encontró manual_usuario.html\n\n{}".format(manual_path),
                    u"Manual no encontrado",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning,
                )
        except Exception as ex:
            try:
                msg = unicode(ex)
            except Exception:
                msg = str(ex)
            MessageBox.Show(
                u"No se pudo abrir el manual:\n{}".format(msg),
                u"Armadura en vigas",
                MessageBoxButton.OK,
                MessageBoxImage.Warning,
            )

    def show(self):
        uidoc = self._revit.ActiveUIDocument
        hwnd = None
        try:
            from System.Windows.Interop import WindowInteropHelper

            hwnd = revit_main_hwnd(self._revit.Application)
            if hwnd:
                helper = WindowInteropHelper(self._win)
                helper.Owner = hwnd
        except Exception:
            pass
        position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        self._document = self._revit.ActiveUIDocument.Document
        self._win.Show()
        self._win.Activate()


def run_pyrevit(revit, manual_dir=None):
    doc = revit.ActiveUIDocument.Document
    rows = _collect_bar_types(doc)
    if not rows:
        MessageBox.Show(
            u"No hay tipos de barra (RebarBarType) en el proyecto.",
            u"Armadura en vigas",
            MessageBoxButton.OK,
            MessageBoxImage.Warning,
        )
        return
    w = ArmaduraVigasCapasWindow(revit, manual_dir, rows)
    w.show()


def run(document, uidocument, manual_dir=None):
    """Entrada alternativa (tests / import)."""
    class _R(object):
        pass

    r = _R()
    r.ActiveUIDocument = uidocument
    run_pyrevit(r, manual_dir)
