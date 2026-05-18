# -*- coding: utf-8 -*-
"""
Revisiones — equivalencia práctica del flujo Dynamo siguienteRevision_script (Emisión).

Por defecto avanza a la revisión inmediatamente siguiente en el orden del proyecto
(`Revision.GetAllRevisionIds`) según la última revisión del índice de cada lámina; no crea nuevas ``Revision``.
Si la lámina no tiene ninguna revisión en ese índice antes de ejecutar la herramienta: en modo automático se usa la primera revisión del proyecto (índice 0 en ``GetAllRevisionIds``); en modo **revisión número 0** se usa la fila cuyo número mostrado es 0 (puede no coincidir con ese índice).
Opcionalmente se puede forzar la **revisión número 0** del proyecto para las láminas seleccionadas; si el índice de la lámina aún no llegó a esa revisión, se añaden las intermedias y se rellena la siguiente fila libre del cajetín (Rnn).
Si la lámina **ya tiene como última revisión del índice** la revisión 0 del proyecto, **no se aplica** emisión en ese modo.
En modo automático (**siguiente revisión**), el índice de Revit avanza como antes (siguiente correlativo en el proyecto);
los datos del formulario se escriben sólo en la **primera fila con NUM vacío** del cajetín (misma regla que revisión 0).

Hay dos convenciones de nombres: ``Rmm_01_NUM``… (fila ``mm``) y ``R01_mm_NUM``… (todas las filas bajo ``R01``); se detecta por los parámetros del cajetín.

Cuando la nueva revisión queda más allá de la posición R20 en el índice de Revit,
el cajetín sigue usando sólo la **primera fila con NUM vacío** y el número de la revisión emitida
(mismo criterio que con índices cortos); puede mostrarse aviso si la posición del índice no coincide con esa fila.

Barra superior de progreso: ``pyrevit.forms.ProgressBar`` mismo estilo que Exportar láminas (título ``Arainco - …``, acento #5BC0DE).
Toda la escritura en lámina(s) se agrupa en **una sola transacción** de Revit (nombre ``Arainco: Revisiones``); si tras abrirla falla una lámina, se revierte el lote entero hasta ese momento.
Ver pushbutton para limitaciones vs Dynamo (Archi-lab, Data-Shapes).
"""

from __future__ import print_function

import codecs
import json
import os
import sys

try:
    unicode
except NameError:
    unicode = str

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Data")

from System import AppDomain, Boolean, DateTime, EventHandler, Int32, Int64, String
from System.Collections.Generic import List as ClrList
from System.Collections.ObjectModel import ObservableCollection
from System.IO import Directory
from System.Data import DataColumn, DataTable
from System.Globalization import CultureInfo
from System.Windows.Markup import XamlReader
from System.Windows import Duration, RoutedEventHandler, WindowState

clr.AddReference("RevitAPI")

clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Revision,
    Transaction,
    ViewSheet,
)

try:
    from Autodesk.Revit.DB import RevisionCloud
except Exception:
    RevisionCloud = None

from Autodesk.Revit.UI import TaskDialog

from pyrevit import forms


APP_DOMAIN_SINGLETON_KEY = "BIMTools.SiguienteRevision.ActiveWindow"


_script_dir_scripts = os.path.dirname(os.path.abspath(__file__))
if _script_dir_scripts not in sys.path:
    sys.path.insert(0, _script_dir_scripts)


_SIGREV_CLOSE_GUARD = {"busy": False, "finalized": False}


def _sigrev_reset_close_guard():
    global _SIGREV_CLOSE_GUARD
    _SIGREV_CLOSE_GUARD = {"busy": False, "finalized": False}


try:
    from exportar_laminas_pdf_dwg import _sheet_revision_display
except Exception:
    _sheet_revision_display = None

from revit_wpf_window_position import revit_main_hwnd  # noqa: E402

try:
    from gestionar_personas_wpf import (  # noqa: E402
        GestionarPersonasDialog,
        load_personas_list,
    )
except Exception:
    GestionarPersonasDialog = None
    load_personas_list = None

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""

try:
    from join_geometry_concrete_vista import (
        _BloquearComandosRevit,
        _pbar_exit_safe as _sigrev_pb_exit_safe,
    )
except Exception:
    _BloquearComandosRevit = None
    _sigrev_pb_exit_safe = None


if _sigrev_pb_exit_safe is None:

    def _sigrev_pb_exit_safe(pb, ok):
        if ok and pb is not None:
            try:
                pb.__exit__(None, None, None)
            except Exception:
                pass


# Misma línea visual que Exportar láminas (pyRevit ProgressBar, acento cian).
_SIGREV_PYREVIT_ACCENT_RGB = (91, 192, 222)
_SIGREV_PBAR_TITLE_BASE = u"Arainco - Revisiones"


def _sigrev_pbar_initial_title(base, total):
    try:
        t = int(total)
    except Exception:
        t = 0
    if t < 1:
        t = 1
    return u"{} 0/{}".format(base, t)


def _sigrev_pbar_step(pb, current_index, count, base_title):
    """Título ``base_title X/Y`` con un solo espacio (mismo formato que Exportar láminas PDF/DWG)."""
    if pb is None:
        return
    c = int(count) if count else 0
    if c < 1:
        c = 1
    i = int(current_index) + 1
    try:
        if hasattr(pb, u"update_progress"):
            try:
                pb.update_progress(i, max_value=c)
            except TypeError:
                try:
                    pb.update_progress(i, max=c)
                except Exception:
                    pass
    except Exception:
        pass
    try:
        pb.title = u"{} {}/{}".format(base_title, i, c)
    except Exception:
        pass


def _sigrev_pbar_start(title, count):
    if count is None or int(count) < 1:
        return None
    try:
        pb = forms.ProgressBar(title=title, cancellable=False)
        try:
            from System.Windows.Media import Color, SolidColorBrush

            r, g, b = _SIGREV_PYREVIT_ACCENT_RGB
            pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(Color.FromRgb(r, g, b))
        except Exception:
            pass
        return pb
    except Exception:
        return None


_SIGREV_CHROME_MS = 260
_SIGREV_WPF_STORYBOARD_DUR = u"0:0:{0:.3f}".format(_SIGREV_CHROME_MS / 1000.0)
# Fallback si Completed del Storyboard no dispara en IronPython + Revit.
_SIGREV_EXIT_FALLBACK_MS = max(380, _SIGREV_CHROME_MS + 120)
_SIGREV_ENTER_FALLBACK_MS = max(400, _SIGREV_CHROME_MS + 160)


DESCRIPCIONES = (
    "PRELIMINAR",
    "PARA APROBACION",
    "EMITIDO PARA APROBACION MUNICIPAL",
    "MODIFICACION GENERAL",
    "APROBADO PARA LICITACION",
    "ACTUALIZACION APL",
    "MODIFICA LO INDICADO",
    "APROBADO PARA CONSTRUCCION",
    "PARA MUNICIPALIDAD",
    "APTO PARA LICITACION",
    "PARA LICITACION",
    "INGRESA UNA DESCRIPCION",
    "PARA INGRESO MUNICIPAL",
    "MODIFICA Y AGREGA LO INDICADO",
    "INFORMATIVO",
    "PARA APROBACION",
    "PARA REVISION",
    "ACTUALIZA LICITACION",
)

DIBUJO_INICIALES = (
    "J.N.N.",
    "P.C.C.",
    "A.S.A.",
    "C.M.H.",
    "S.J.M.",
    "R.P.C.",
    "C.O.C.",
    "H.C.P.",
    "L.B.M.",
    "J.N.O.",
    "T.M.M.",
    "B.L.A.",
    "X.X.X.",
)

# Misma base que 01_BIMIssue / 03_Personas (directorio compartido personas.json).
ISSUES_DIR = u"Y:\\00_SERVIDOR DE INCIDENCIAS"
PERSONAS_FILE = os.path.join(ISSUES_DIR, "personas.json")

PERSONA_ROL_MODELADOR = u"Modelador"
PERSONA_ROL_INGENIERO = u"Ingeniero"


def _normalize_persona_rol_sigrev(val):
    s = (val or u"").strip()
    if s == PERSONA_ROL_INGENIERO:
        return PERSONA_ROL_INGENIERO
    return PERSONA_ROL_MODELADOR


def _sigrev_personas_display_map_for_rol(target_rol):
    """
    personas.json filtradas por rol: texto del combo = nombre completo (o abreviación si no hay nombre);
    valor para escribir en lámina = abreviación si existe, si no el nombre.
    """
    items_order = []
    display_to_sheet = {}
    if target_rol not in (PERSONA_ROL_MODELADOR, PERSONA_ROL_INGENIERO):
        return items_order, display_to_sheet
    if not os.path.isfile(PERSONAS_FILE):
        return items_order, display_to_sheet
    try:
        with codecs.open(PERSONAS_FILE, "r", "utf-8-sig") as f:
            data = json.load(f)
    except Exception:
        return items_order, display_to_sheet
    if not isinstance(data, list):
        return items_order, display_to_sheet
    seen = set()
    for p in data:
        if not isinstance(p, dict):
            continue
        if _normalize_persona_rol_sigrev(p.get("rol", u"")) != target_rol:
            continue
        abr = (p.get("abreviacion") or u"").strip()
        nom = (p.get("nombre") or u"").strip()
        display = nom if nom else abr
        if not display:
            continue
        sheet_val = abr if abr else nom
        key = display.lower()
        if key in seen:
            continue
        seen.add(key)
        items_order.append(display)
        display_to_sheet[display] = sheet_val
    try:
        items_order.sort(key=lambda s: s.lower())
    except Exception:
        items_order.sort()
    return items_order, display_to_sheet


def _sigrev_combo_display_to_sheet_value(selected_display, mapping):
    """Convierte el ítem seleccionado del combo (nombre mostrado) al texto para parámetros de lámina."""
    label = unicode(selected_display or u"").strip()
    if not label:
        return u""
    m = mapping or {}
    return unicode(m.get(label, label)).strip()


def _sigrev_fill_persona_combos(win, state):
    """
    Rellena Dibujó (modeladores), Revisó y Aprobó (ingenieros). En el formulario se muestra el nombre;
    state['sigrev_map_dib'] / state['sigrev_map_ing'] traducen a abreviación para las láminas.
    """
    cb_d = win.FindName("CbDibujo")
    cb_r = win.FindName("CbReviso")
    cb_a = win.FindName("CbAprobo")
    if cb_d is None or cb_r is None or cb_a is None:
        return
    cb_d.Items.Clear()
    cb_r.Items.Clear()
    cb_a.Items.Clear()
    state["sigrev_map_dib"] = {}
    state["sigrev_map_ing"] = {}
    dibujo_items, dib_map = _sigrev_personas_display_map_for_rol(PERSONA_ROL_MODELADOR)
    if not dibujo_items:
        dibujo_items = list(DIBUJO_INICIALES)
        dib_map = {x: x for x in dibujo_items}
    ing_items, ing_map = _sigrev_personas_display_map_for_rol(PERSONA_ROL_INGENIERO)
    if not ing_items:
        ing_items = list(DIBUJO_INICIALES)
        ing_map = {x: x for x in ing_items}
    state["sigrev_map_dib"] = dib_map
    state["sigrev_map_ing"] = ing_map
    for x in dibujo_items:
        cb_d.Items.Add(x)
    for x in ing_items:
        cb_r.Items.Add(x)
        cb_a.Items.Add(x)
    cb_d.SelectedIndex = 0
    cb_r.SelectedIndex = 0
    cb_a.SelectedIndex = 0


# Bloques de emisión en instancia de lámina (mismo convención que listado_planos_excel_core).
MAX_REVISION_SLOTS = 20
RNN_SUFFIX_NUM = u"01_NUM"
RNN_SUFFIX_DES = u"02_DES"
RNN_SUFFIX_DIR = u"03_DIR"
RNN_SUFFIX_DIB = u"03_DIB"  # mismo campo «dibujó» en familias que no usan DIR
RNN_SUFFIX_REV = u"04_REV"
RNN_SUFFIX_APR = u"05_APR"
RNN_SUFFIX_FCH = u"06_FCH"

# Parámetro en instancias de nube de revisión (opcional); si el nombre no coincide, se ignora.
PARAM_CANTIDAD_REVISIONES = u"CANTIDAD_REVISIONES"


REV_XAML = (
    u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    x:Name="SigRevWin"
    Title="Revisiones"
    Height="1060" Width="1040" MinHeight="920" MinWidth="800"
    Background="Transparent"
    AllowsTransparency="True"
    WindowStyle="None"
    ResizeMode="NoResize"
    WindowStartupLocation="CenterScreen"
    Topmost="True"
    Opacity="1"
    UseLayoutRounding="True"
    FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
    <!-- Revisiones: animación solo sobre SigRevAnimShell (interior). SigRevRootChrome permanece opaco (evita halo vacío en capturas). -->
    <Storyboard x:Key="SigRevEnterStoryboard">
      <DoubleAnimation Storyboard.TargetName="SigRevAnimShell" Storyboard.TargetProperty="Opacity"
                       From="0" To="1" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="SigRevEnterTranslate" Storyboard.TargetProperty="Y"
                       From="22" To="0" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
    </Storyboard>
    <Storyboard x:Key="SigRevExitStoryboard">
      <DoubleAnimation Storyboard.TargetName="SigRevAnimShell" Storyboard.TargetProperty="Opacity"
                       From="1" To="0" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseIn"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="SigRevEnterTranslate" Storyboard.TargetProperty="Y"
                       From="0" To="22" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseIn"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
    </Storyboard>
"""
    + BIMTOOLS_DARK_STYLES_XML
    + u"""
    <Style x:Key="SigRevComboItem" TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}">
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="FontSize" Value="13"/>
    </Style>
    <Style x:Key="SigRevComboStretch" TargetType="ComboBox" BasedOn="{StaticResource Combo}">
      <Setter Property="Width" Value="{x:Static sys:Double.NaN}"/>
      <Setter Property="Height" Value="{x:Static sys:Double.NaN}"/>
      <Setter Property="HorizontalAlignment" Value="Stretch"/>
      <Setter Property="MinWidth" Value="0"/>
      <Setter Property="MaxWidth" Value="99999"/>
      <Setter Property="MinHeight" Value="34"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Foreground" Value="#F2F8FC"/>
      <Setter Property="Margin" Value="0,0,0,0"/>
      <Setter Property="ItemContainerStyle" Value="{StaticResource SigRevComboItem}"/>
    </Style>
    <!-- Misma geometría en columnas pareadas del formulario de emisión (sin anchos fijos). -->
    <Style x:Key="SigRevComboHalf" TargetType="ComboBox" BasedOn="{StaticResource SigRevComboStretch}">
      <Setter Property="Margin" Value="0,0,0,0"/>
    </Style>
    <Style x:Key="StepBadge" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#5BC0DE"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>
    <Style x:Key="SigRevFormLabel" TargetType="TextBlock" BasedOn="{StaticResource Label}">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
    </Style>
    <Style x:Key="SigRevStepBadge" TargetType="TextBlock" BasedOn="{StaticResource StepBadge}">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
    </Style>
    <Style x:Key="SigRevRadio" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,8,0,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="PanelInset" TargetType="Border">
      <Setter Property="Background" Value="#071018"/>
      <Setter Property="BorderBrush" Value="#1E3F55"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="12,10"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="BtnGhost" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#9BC4D6"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Padding" Value="12,7"/>
      <Setter Property="BorderBrush" Value="#2A4A5E"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="R" Background="{TemplateBinding Background}" CornerRadius="5"
                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="R" Property="Background" Value="#0D1E2E"/>
                <Setter TargetName="R" Property="BorderBrush" Value="#5BC0DE"/>
                <Setter Property="Foreground" Value="#E8F4F8"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="DataGrid" BasedOn="{StaticResource {x:Type DataGrid}}">
      <Setter Property="Background" Value="#040A12"/>
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="RowBackground" Value="#0B1726"/>
      <Setter Property="AlternatingRowBackground" Value="#071420"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#152A3D"/>
      <Setter Property="VerticalGridLinesBrush" Value="#152A3D"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="RowHeight" Value="34"/>
      <Setter Property="GridLinesVisibility" Value="All"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource {x:Type DataGridColumnHeader}}">
      <Setter Property="Background" Value="#0F2840"/>
      <Setter Property="Foreground" Value="#C8E4EF"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Padding" Value="12,10"/>
      <Setter Property="BorderBrush" Value="#1A3A50"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>
    <Style TargetType="DataGridRow" BasedOn="{StaticResource {x:Type DataGridRow}}">
      <Setter Property="Background" Value="#0B1726"/>
      <Style.Triggers>
        <Trigger Property="AlternationIndex" Value="0">
          <Setter Property="Background" Value="#0B1726"/>
        </Trigger>
        <Trigger Property="AlternationIndex" Value="1">
          <Setter Property="Background" Value="#071420"/>
        </Trigger>
        <MultiTrigger>
          <MultiTrigger.Conditions>
            <Condition Property="IsSelected" Value="False"/>
            <Condition Property="IsMouseOver" Value="True"/>
          </MultiTrigger.Conditions>
          <Setter Property="Background" Value="#132A40"/>
        </MultiTrigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#1B3D5C"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="GridCellPadding" TargetType="DataGridCell">
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="Transparent"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="ChkSheetSel" TargetType="CheckBox">
      <Setter Property="Foreground" Value="#C8E4EF"/>
      <Setter Property="HorizontalAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Width" Value="18"/>
      <Setter Property="Height" Value="18"/>
    </Style>
    <Style x:Key="DgTbLeft" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextAlignment" Value="Left"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="DgTbLeftPadded" TargetType="TextBlock" BasedOn="{StaticResource DgTbLeft}">
      <Setter Property="Margin" Value="12,0,8,0"/>
    </Style>
    <Style x:Key="DgTbCenter" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#050E18"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#284760"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="FontSize" Value="11"/>
    </Style>
    <Style x:Key="DgHdrCenter" TargetType="DataGridColumnHeader" BasedOn="{StaticResource {x:Type DataGridColumnHeader}}">
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="DgHdrLeft" TargetType="DataGridColumnHeader" BasedOn="{StaticResource {x:Type DataGridColumnHeader}}">
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
    </Style>
    <Style x:Key="ExpLamScrollBarDark" TargetType="ScrollBar" BasedOn="{StaticResource BimToolsScrollBarDark}"/>
  </Window.Resources>
  <Border x:Name="SigRevRootChrome" Opacity="1" CornerRadius="8" Background="#0E1B32" Padding="14"
          BorderBrush="#5BC0DE" BorderThickness="1" ClipToBounds="True">
    <Border x:Name="SigRevAnimShell" Background="Transparent" Opacity="0">
      <Border.RenderTransform>
        <TranslateTransform x:Name="SigRevEnterTranslate" Y="22"/>
      </Border.RenderTransform>
      <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <Border x:Name="TitleBar" Grid.Row="0" Background="#0D1E2E" CornerRadius="6" Padding="12,10" Margin="0,0,0,10"
              BorderBrush="#5BC0DE" BorderThickness="1" HorizontalAlignment="Stretch">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center"
                      ToolTip="Emisión en láminas: nueva entrada de revisión por lámina seleccionada.">
            <Image x:Name="ImgLogo" Width="40" Height="40"
                   Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center"
                   RenderOptions.BitmapScalingMode="HighQuality"/>
            <TextBlock Text="Revisiones" FontSize="14" FontWeight="Bold"
                       Foreground="#FFFFFF" TextWrapping="NoWrap" VerticalAlignment="Center"
                       TextAlignment="Left"/>
          </StackPanel>
          <Button x:Name="BtnClose" Grid.Column="2"
                  Style="{StaticResource BtnCloseX_MinimalNoBg}"
                  VerticalAlignment="Center" HorizontalAlignment="Right" ToolTip="Cerrar"/>
        </Grid>
      </Border>
      <Border Grid.Row="2" Style="{StaticResource PanelInset}" Margin="0,0,0,10" Padding="10,8">
        <StackPanel>
          <Grid VerticalAlignment="Center">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Margin="0,0,10,0" Background="#050E18" BorderBrush="#355973" BorderThickness="1" CornerRadius="5" Padding="0" MinHeight="32">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="&#xE721;" FontFamily="Segoe MDL2 Assets" FontSize="15"
                           Foreground="#6B94AA" VerticalAlignment="Center" Margin="10,0,4,0" IsHitTestVisible="False"/>
                <Grid Grid.Column="1" MinHeight="28">
                  <TextBox x:Name="TxtBuscar" Background="Transparent" BorderThickness="0"
                           VerticalContentAlignment="Center" Padding="0,6,10,6"
                           ToolTip="Filtrar por número, nombre, revisión actual o nueva revisión"/>
                  <TextBlock x:Name="TxtBuscarWatermark" Text="Buscar" IsHitTestVisible="False"
                             Foreground="#5C7A8F" FontSize="12" VerticalAlignment="Center" Margin="0,0,10,0"/>
                </Grid>
              </Grid>
            </Border>
            <Button x:Name="BtnRefrescar" Grid.Column="1" Style="{StaticResource BtnGhost}" Padding="8,5" FontSize="10"
                    ToolTip="Actualizar la lista desde el proyecto">
              <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="&#xE72C;" FontFamily="Segoe MDL2 Assets" FontSize="13"
                           VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBlock Text="Actualizar" VerticalAlignment="Center" FontSize="10" FontWeight="SemiBold"/>
              </StackPanel>
            </Button>
          </Grid>
        </StackPanel>
      </Border>
      <Border Grid.Row="1" Style="{StaticResource PanelInset}" Margin="0,0,0,10" Padding="14,14">
        <!-- Descripción ancha · parejas Dibujó/Revisó y Aprobó/Fecha · modo revisión. -->
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="22"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,0,0,14">
            <TextBlock Text="Descripción" Style="{StaticResource SigRevFormLabel}"/>
            <ComboBox x:Name="CbDescripcion" Style="{StaticResource SigRevComboStretch}"/>
          </StackPanel>
          <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,0,0,14">
            <TextBlock Text="Dibujó" Style="{StaticResource SigRevFormLabel}"/>
            <ComboBox x:Name="CbDibujo" Style="{StaticResource SigRevComboHalf}"/>
          </StackPanel>
          <StackPanel Grid.Row="1" Grid.Column="2" Margin="0,0,0,14">
            <TextBlock Text="Revisó" Style="{StaticResource SigRevFormLabel}"/>
            <ComboBox x:Name="CbReviso" Style="{StaticResource SigRevComboHalf}"/>
          </StackPanel>
          <StackPanel Grid.Row="2" Grid.Column="0" Margin="0,0,0,14">
            <TextBlock Text="Aprobó" Style="{StaticResource SigRevFormLabel}"/>
            <ComboBox x:Name="CbAprobo" Style="{StaticResource SigRevComboHalf}"/>
          </StackPanel>
          <StackPanel Grid.Row="2" Grid.Column="2" Margin="0,0,0,14">
            <TextBlock Text="Fecha (dd.MM.yy)" Style="{StaticResource SigRevFormLabel}"/>
            <ComboBox x:Name="CbFecha" Style="{StaticResource SigRevComboHalf}"/>
          </StackPanel>
          <Button x:Name="BtnGestionarPersonas" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="3"
                  Content="Gestionar directorio de personas..." HorizontalAlignment="Left"
                  Style="{StaticResource BtnGhost}" Margin="0,0,0,12" Padding="12,7" FontSize="11"
                  ToolTip="Dar de alta o editar personas en personas.json (modeladores e ingenieros)."/>
          <StackPanel Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,4,0,4">
            <TextBlock Text="Modo de Revisión" Style="{StaticResource SigRevStepBadge}"/>
            <RadioButton x:Name="RadRevAutomatica" GroupName="SigRevDestino" Style="{StaticResource SigRevRadio}"
                         Content="Subir a siguiente revisión." IsChecked="True" Margin="0,0,0,0"/>
            <RadioButton x:Name="RadRevisionPuntual" GroupName="SigRevDestino" Style="{StaticResource SigRevRadio}"
                         Content="Subir a Revisión 0." Margin="0,8,0,0"
                         ToolTip="Usa la revisión cuyo número en Gestión de revisiones es 0. Si el índice de la lámina aún no la incluye, se añaden las intermedias y se rellena la siguiente fila Rnn vacía del cajetín. Las láminas que ya tienen la revisión 0 como última del índice no se modifican."/>
          </StackPanel>
        </Grid>
      </Border>
      <Border Grid.Row="3" Background="#040A12" BorderBrush="#1E3F55" BorderThickness="1" CornerRadius="8" Padding="0" Margin="0,0,0,12">
        <DataGrid x:Name="GridSheets" MinHeight="396"
                  AutoGenerateColumns="False" CanUserAddRows="False" CanUserResizeRows="False"
                  RowHeaderWidth="0" SelectionMode="Extended" SelectionUnit="FullRow"
                  AlternationCount="2" ClipboardCopyMode="ExcludeHeader" HeadersVisibility="Column"
                  VerticalAlignment="Stretch" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <DataGrid.CellStyle>
            <Style TargetType="DataGridCell" BasedOn="{StaticResource GridCellPadding}"/>
          </DataGrid.CellStyle>
          <DataGrid.Columns>
            <DataGridCheckBoxColumn Header="" Binding="{Binding Sel, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="44" CanUserSort="False">
              <DataGridCheckBoxColumn.HeaderTemplate>
                <DataTemplate>
                  <CheckBox Tag="HdrSelectAll" Style="{StaticResource ChkSheetSel}" HorizontalAlignment="Center"
                            VerticalAlignment="Center" IsThreeState="True"
                            ToolTip="Marcar o anular todas las filas visibles elegibles (respeta Buscar; en Revisión 0 no aplica a láminas ya en rev. 0)."/>
                </DataTemplate>
              </DataGridCheckBoxColumn.HeaderTemplate>
              <DataGridCheckBoxColumn.HeaderStyle>
                <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrCenter}">
                  <Setter Property="ToolTip" Value="Casilla por fila (deshabilitada si ya está en rev. 0 en modo Revisión 0). Cabecera: marcar/anular todas las elegibles visibles. Mayús+clic: rango."/>
                </Style>
              </DataGridCheckBoxColumn.HeaderStyle>
              <DataGridCheckBoxColumn.ElementStyle>
                <Style TargetType="CheckBox" BasedOn="{StaticResource ChkSheetSel}">
                  <Setter Property="IsEnabled" Value="{Binding SelEnabled}"/>
                </Style>
              </DataGridCheckBoxColumn.ElementStyle>
              <DataGridCheckBoxColumn.EditingElementStyle>
                <Style TargetType="CheckBox" BasedOn="{StaticResource ChkSheetSel}">
                  <Setter Property="IsEnabled" Value="{Binding SelEnabled}"/>
                </Style>
              </DataGridCheckBoxColumn.EditingElementStyle>
            </DataGridCheckBoxColumn>
            <DataGridTextColumn Header="Número" Binding="{Binding SheetNumber}" Width="112" IsReadOnly="True">
              <DataGridTextColumn.HeaderStyle>
                <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrCenter}"/>
              </DataGridTextColumn.HeaderStyle>
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbCenter}"/>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="Nombre" Binding="{Binding SheetName}" Width="*" MinWidth="220" IsReadOnly="True">
              <DataGridTextColumn.HeaderStyle>
                <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrLeft}"/>
              </DataGridTextColumn.HeaderStyle>
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbLeftPadded}"/>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="Revisión actual" Binding="{Binding Revision}" Width="118" MinWidth="100" IsReadOnly="True">
              <DataGridTextColumn.HeaderStyle>
                <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrCenter}"/>
              </DataGridTextColumn.HeaderStyle>
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbCenter}"/>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="Nueva revisión" Binding="{Binding NuevaRevision}" Width="108" MinWidth="92" IsReadOnly="True">
              <DataGridTextColumn.HeaderStyle>
                <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrCenter}"/>
              </DataGridTextColumn.HeaderStyle>
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbCenter}"/>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="" Binding="{Binding IdInt}" Width="0" MinWidth="0" MaxWidth="0" IsReadOnly="True">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbLeft}"/>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
          </DataGrid.Columns>
        </DataGrid>
      </Border>
      <Border Grid.Row="4" Background="#050E18" BorderBrush="#1E3F55" BorderThickness="1" CornerRadius="6"
              Padding="14,12">
        <Border.Effect>
          <DropShadowEffect BlurRadius="18" ShadowDepth="0" Opacity="0.2" Color="#000000"/>
        </Border.Effect>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0" VerticalAlignment="Center" Orientation="Vertical"
                      ToolTip="Columnas Revisión actual / Nueva revisión: última en índice vs la que se emitiría con el modo elegido. Primera columna: incluir en la emisión.">
            <TextBlock x:Name="TxtEstadoLaminas" Foreground="#E8F4F8" FontSize="12" FontWeight="SemiBold"
                       TextWrapping="Wrap" Text="0 láminas  |  0 seleccionadas"/>
            <TextBlock Foreground="#6B94AA" FontSize="10" Margin="0,4,0,0"
                       Text="Estado del listado · listo para emitir"/>
          </StackPanel>
          <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
            <Button x:Name="BtnCancel" Content="Cancelar" Style="{StaticResource BtnGhost}" MinWidth="100" Margin="0,0,10,0"/>
            <Button x:Name="BtnOk" Content="Iniciar" Style="{StaticResource BtnPrimary}" MinWidth="132" MinHeight="38"/>
          </StackPanel>
        </Grid>
      </Border>
    </Grid>
    </Border>
  </Border>
</Window>
"""
).replace(u"__WPF_STORYBOARD_DUR__", _SIGREV_WPF_STORYBOARD_DUR)


def _sigrev_load_logo(win):
    try:
        import bimtools_paths

        img = win.FindName(u"ImgLogo")
        if img is None:
            return
        bmp = bimtools_paths.load_logo_bitmap_image()
        if bmp is None:
            return
        img.Source = bmp
        try:
            win.Icon = bmp
        except Exception:
            pass
    except Exception:
        pass


def _sigrev_apply_scrollbars_below(root_visual, resources_owner):
    try:
        from System.Windows.Controls.Primitives import ScrollBar
        from System.Windows.Media import VisualTreeHelper
        from System.Windows import FrameworkElement

        sb_type = clr.GetClrType(ScrollBar)
        st = None
        if resources_owner is not None:
            for _key in (u"ExpLamScrollBarDark", u"BimToolsScrollBarDark"):
                try:
                    st = resources_owner.TryFindResource(_key)
                except Exception:
                    st = None
                if st is not None:
                    break
            if st is None:
                try:
                    st = resources_owner.TryFindResource(sb_type)
                except Exception:
                    st = None
        if st is None:
            return

        def _walk(o, depth):
            if depth > 60 or o is None:
                return
            try:
                n = VisualTreeHelper.GetChildrenCount(o)
                for i in range(n):
                    ch = VisualTreeHelper.GetChild(o, i)
                    try:
                        if ch.GetType().Equals(sb_type):
                            try:
                                ch.ClearValue(FrameworkElement.StyleProperty)
                            except Exception:
                                pass
                            ch.Style = st
                    except Exception:
                        pass
                    _walk(ch, depth + 1)
            except Exception:
                pass

        _walk(root_visual, 0)
    except Exception:
        pass


# --- Animación de ventana (solo Revisiones / REV_XAML): SigRevEnterStoryboard, SigRevExitStoryboard, SigRevAnimShell ---
def _sigrev_force_outer_chrome_visible(win):
    """Opacidad del marco exterior y de la ventana; no modifica SigRevAnimShell (entrada animada)."""
    try:
        ch = win.FindName("SigRevRootChrome")
        if ch is not None:
            ch.Opacity = 1.0
        win.Opacity = 1.0
    except Exception:
        pass


def _sigrev_snap_anim_shell_to_visible(win):
    """Shell interior visible y sin animaciones pending (fallback IronPython / antes del fade de salida)."""
    try:
        from System.Windows import UIElement
        from System.Windows.Media import TranslateTransform

        shell = win.FindName("SigRevAnimShell")
        tt = win.FindName("SigRevEnterTranslate")
        if shell is not None:
            try:
                shell.BeginAnimation(UIElement.OpacityProperty, None)
            except Exception:
                pass
            shell.Opacity = 1.0
        if tt is not None:
            try:
                tt.BeginAnimation(TranslateTransform.YProperty, None)
            except Exception:
                pass
            tt.Y = 0.0
    except Exception:
        pass


def _sigrev_schedule_scrollbar_styles(win):
    try:
        from System import Action
        from System.Windows.Threading import DispatcherPriority

        def _go():
            _sigrev_apply_scrollbars_below(win, win)

        _go()
        win.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, Action(_go))
    except Exception:
        try:
            _sigrev_apply_scrollbars_below(win, win)
        except Exception:
            pass


def _sigrev_begin_open_storyboard(win):
    """
    Fade-in + slide-up solo sobre SigRevAnimShell (IronPython + ShowDialog).
    Reserva de seguridad: Completed + DispatcherTimer llaman a _sigrev_snap_anim_shell_to_visible.
    """
    if getattr(win, "_sigrev_open_sb_done", False):
        return
    win._sigrev_open_sb_done = True
    try:
        win.Opacity = 1.0
    except Exception:
        pass
    _sigrev_force_outer_chrome_visible(win)

    try:
        from System import TimeSpan
        from System.Windows import UIElement
        from System.Windows.Media import TranslateTransform
        from System.Windows.Threading import DispatcherTimer

        shell = win.FindName("SigRevAnimShell")
        tt = win.FindName("SigRevEnterTranslate")
        sb = win.TryFindResource("SigRevEnterStoryboard")
        if shell is None or tt is None or sb is None:
            _sigrev_snap_anim_shell_to_visible(win)
            return

        try:
            shell.BeginAnimation(UIElement.OpacityProperty, None)
            tt.BeginAnimation(TranslateTransform.YProperty, None)
        except Exception:
            pass
        shell.Opacity = 0.0
        tt.Y = 22.0

        dur = Duration(TimeSpan.FromMilliseconds(float(_SIGREV_CHROME_MS)))
        try:
            for i in range(int(sb.Children.Count)):
                sb.Children[i].Duration = dur
        except Exception:
            pass

        win._sigrev_enter_anim_done = False

        def _enter_finish(_s=None, _e=None):
            if getattr(win, "_sigrev_enter_anim_done", False):
                return
            win._sigrev_enter_anim_done = True
            try:
                tm = getattr(win, "_sigrev_enter_fallback_timer", None)
                if tm is not None:
                    try:
                        tm.Stop()
                    except Exception:
                        pass
            except Exception:
                pass
            _sigrev_snap_anim_shell_to_visible(win)

        sb.Completed += EventHandler(_enter_finish)
        sb.Begin(win, True)

        try:
            tm = DispatcherTimer()
            tm.Interval = TimeSpan.FromMilliseconds(float(_SIGREV_ENTER_FALLBACK_MS))
            tm.Tick += EventHandler(lambda _snd, _evt: _enter_finish())
            win._sigrev_enter_fallback_timer = tm
            tm.Start()
        except Exception:
            pass
    except Exception:
        _sigrev_snap_anim_shell_to_visible(win)


def _sigrev_on_win_loaded(win):
    _sigrev_force_outer_chrome_visible(win)
    try:
        _sigrev_schedule_scrollbar_styles(win)
    except Exception:
        pass
    try:
        gd = win.FindName("GridSheets")
        if gd is not None:
            _sigrev_schedule_sheet_grid_scrollbars(win, gd)
    except Exception:
        pass
    try:
        _sigrev_begin_open_storyboard(win)
    except Exception:
        _sigrev_force_outer_chrome_visible(win)
        _sigrev_snap_anim_shell_to_visible(win)


def _sigrev_schedule_sheet_grid_scrollbars(win, dg):
    try:
        from System import Action
        from System.Windows.Threading import DispatcherPriority

        def _go():
            _sigrev_apply_scrollbars_below(dg, win)
            _sigrev_apply_scrollbars_below(win, win)

        _go()
        win.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, Action(_go))
    except Exception:
        try:
            _sigrev_apply_scrollbars_below(dg, win)
            _sigrev_apply_scrollbars_below(win, win)
        except Exception:
            pass


def _collect_sheets(doc):
    out = []
    for vs in FilteredElementCollector(doc).OfClass(ViewSheet):
        if getattr(vs, "IsPlaceholder", False):
            continue
        out.append(vs)
    return sorted(out, key=lambda s: (s.SheetNumber or "", s.Name or ""))


def _sheet_display(sheet):
    return u"{} - {}".format(sheet.SheetNumber or "?", sheet.Name or "?")


def _sigrev_sheet_revision_cell(sheet, doc):
    if _sheet_revision_display is None:
        return u""
    try:
        return (unicode(_sheet_revision_display(sheet, doc)).strip())
    except Exception:
        return u""

def _sigrev_row_id_for_table(sheet):
    """
    Valor numérico estable de sheet.Id para persistir en DataTable.
    Revit 2026+ usa identificadores 64 bits (ElementId.Value); no usar Int32 en la tabla.
    """
    return int(_element_id_integer(sheet.Id))


# Láminas con esta subcadena en el nombre no aparecen en el listado de emisión.
_SIGREV_EXCLUDE_SHEET_NAME_SUBSTR = u"splash screen"


def _build_sheet_selection_table(doc, sheets_all):
    """
    Sel, SelEnabled, SheetNumber, SheetName, Revision, NuevaRevision, IdInt.

    SelEnabled: en modo «Revisión 0», False si la lámina ya tiene rev. 0 como última en índice
    (no se permite marcar el checkbox).

    Omite láminas cuyo nombre contiene ``Splash Screen`` (insensible a mayúsculas).
    """
    tbl = DataTable()
    tbl.Columns.Add(DataColumn(u"Sel", clr.GetClrType(Boolean)))
    tbl.Columns.Add(DataColumn(u"SelEnabled", clr.GetClrType(Boolean)))
    tbl.Columns.Add(DataColumn(u"SheetNumber", clr.GetClrType(String)))
    tbl.Columns.Add(DataColumn(u"SheetName", clr.GetClrType(String)))
    tbl.Columns.Add(DataColumn(u"Revision", clr.GetClrType(String)))
    tbl.Columns.Add(DataColumn(u"NuevaRevision", clr.GetClrType(String)))
    tbl.Columns.Add(DataColumn(u"IdInt", clr.GetClrType(Int64)))
    try:
        tbl.BeginLoadData()
    except Exception:
        pass
    for s in sheets_all:
        nm = unicode(s.Name or u"")
        if _SIGREV_EXCLUDE_SHEET_NAME_SUBSTR in nm.lower():
            continue
        row = tbl.NewRow()
        row[u"Sel"] = False
        row[u"SelEnabled"] = True
        row[u"SheetNumber"] = unicode(s.SheetNumber or u"")
        row[u"SheetName"] = unicode(s.Name or u"")
        row[u"Revision"] = _sigrev_sheet_revision_cell(s, doc)
        row[u"NuevaRevision"] = u""
        row[u"IdInt"] = Int64(_sigrev_row_id_for_table(s))
        tbl.Rows.Add(row)
    try:
        tbl.EndLoadData()
    except Exception:
        pass
    try:
        tbl.AcceptChanges()
    except Exception:
        pass
    return tbl


def _collect_sheets_checked_in_table(doc, tbl):
    out = []
    if tbl is None:
        return out
    n = int(tbl.Rows.Count)
    for i in range(n):
        row = tbl.Rows[i]
        try:
            if not bool(row[u"Sel"]):
                continue
            eid_val = row[u"IdInt"]
            el = doc.GetElement(_sigrev_element_id_from_table_cell(eid_val))
            if isinstance(el, ViewSheet):
                out.append(el)
        except Exception:
            continue
    return out


def _sigrev_element_id_from_table_cell(eid_val):
    """ElementId desde la celda IdInt (Int64); imprescindible en Revit 2026+ (ids 64 bits)."""
    try:
        return ElementId(Int64(int(eid_val)))
    except Exception:
        try:
            return ElementId(int(eid_val))
        except Exception:
            return ElementId.InvalidElementId


def _sigrev_grid_index_of_datarowview(grid, drv):
    """Índice en grid.Items; fallback por IdInt si IndexOf no reconoce el DataContext (p. ej. WPF 2026)."""
    if grid is None or drv is None:
        return -1
    try:
        idx = int(grid.Items.IndexOf(drv))
        if idx >= 0:
            return idx
    except Exception:
        pass
    try:
        want = int(drv.Row[u"IdInt"])
    except Exception:
        try:
            want = int(drv[u"IdInt"])
        except Exception:
            return -1
    try:
        n = int(grid.Items.Count)
    except Exception:
        return -1
    for i in range(n):
        try:
            it = grid.Items[i]
            if int(it.Row[u"IdInt"]) == want:
                return i
        except Exception:
            continue
    return -1


def _sigrev_nullable_bool(wpf_nb):
    try:
        if wpf_nb is None:
            return False
        if hasattr(wpf_nb, "HasValue"):
            return bool(wpf_nb.HasValue and wpf_nb.Value)
        return unicode(wpf_nb).strip().lower() == u"true"
    except Exception:
        return False


def _sigrev_row_selected_rowview(rv):
    if rv is None:
        return False
    try:
        if bool(rv.Row[u"Sel"]):
            return True
    except Exception:
        pass
    try:
        return _sigrev_nullable_bool(rv[u"Sel"])
    except Exception:
        return False


def _sigrev_detach_row_changed(state):
    tbl = state.get("sheet_table")
    d = state.get("_sigrev_row_delegate")
    if tbl is not None and d is not None:
        try:
            tbl.RowChanged -= d
        except Exception:
            pass


def _sigrev_make_row_delegate(state):
    from System.Data import DataRowChangeEventHandler

    def _fn(sender, args):
        _sigrev_on_table_row_changed(state, sender, args)

    return DataRowChangeEventHandler(_fn)


def _sigrev_refresh_estado_sheet(state):
    win = state.get("win")
    tbl = state.get("sheet_table")
    tb = win.FindName("TxtEstadoLaminas") if win is not None else None
    if tbl is None or tb is None:
        return
    n = int(tbl.Rows.Count)
    ns = 0
    for i in range(n):
        try:
            if bool(tbl.Rows[i][u"Sel"]):
                ns += 1
        except Exception:
            pass
    try:
        tb.Text = u"{0} láminas  |  {1} seleccionadas".format(n, ns)
    except Exception:
        pass


def _sigrev_sync_buscar_wm(state):
    win = state.get("win")
    if win is None:
        return
    try:
        from System.Windows import Visibility

        tb = win.FindName("TxtBuscar")
        wm = win.FindName("TxtBuscarWatermark")
        if tb is None or wm is None:
            return
        txt = unicode(tb.Text or u"").strip()
        foc = False
        try:
            foc = bool(tb.IsFocused)
        except Exception:
            pass
        show = not txt and not foc
        wm.Visibility = Visibility.Visible if show else Visibility.Collapsed
    except Exception:
        pass


def _sigrev_refresh_nueva_revision_column(doc, tbl, revision_emit_revision_zero, ti_rev0, ordered=None):
    if tbl is None or doc is None:
        return
    if ordered is None:
        ordered = _revision_ids_ordered_in_project(doc)
    n = int(tbl.Rows.Count)
    for i in range(n):
        row = tbl.Rows[i]
        try:
            eid_val = row[u"IdInt"]
        except Exception:
            try:
                row[u"NuevaRevision"] = u"—"
                row[u"SelEnabled"] = True
            except Exception:
                pass
            continue
        try:
            el = doc.GetElement(_sigrev_element_id_from_table_cell(eid_val))
        except Exception:
            el = None
        if not isinstance(el, ViewSheet):
            try:
                row[u"NuevaRevision"] = u"—"
                row[u"SelEnabled"] = True
            except Exception:
                pass
            continue
        try:
            row[u"NuevaRevision"] = _sigrev_nueva_revision_preview_for_sheet(
                doc, el, ordered, revision_emit_revision_zero, ti_rev0
            )
        except Exception:
            try:
                row[u"NuevaRevision"] = u"—"
            except Exception:
                pass
        try:
            en = _sigrev_compute_sel_enabled_for_sheet(
                doc, el, ordered, revision_emit_revision_zero, ti_rev0
            )
            row[u"SelEnabled"] = bool(en)
            if not en:
                row[u"Sel"] = False
        except Exception:
            try:
                row[u"SelEnabled"] = True
            except Exception:
                pass


def _sigrev_refresh_nueva_revision_from_ui(state):
    doc = state.get("doc")
    tbl = state.get("sheet_table")
    win = state.get("win")
    if doc is None or tbl is None:
        return
    ordered = _revision_ids_ordered_in_project(doc)
    ti_rev0 = _index_revision_project_display_number(doc, ordered, u"0")
    emit_rev0 = False
    if win is not None:
        rp = win.FindName("RadRevisionPuntual")
        if rp is not None:
            emit_rev0 = _sigrev_nullable_bool(rp.IsChecked)
    _sigrev_refresh_nueva_revision_column(doc, tbl, emit_rev0, ti_rev0, ordered)
    _sigrev_refresh_estado_sheet(state)
    _sigrev_sync_sel_all_hdr(state)


def _sigrev_apply_buscar_filter(state):
    win = state.get("win")
    tbl = state.get("sheet_table")
    if win is None or tbl is None:
        return
    dv = tbl.DefaultView
    try:
        t = unicode(win.FindName("TxtBuscar").Text).strip()
    except Exception:
        t = u""
    if not t:
        dv.RowFilter = u""
    else:
        esc = t.replace(u"'", u"''")
        dv.RowFilter = (
            u"[SheetNumber] LIKE '%{0}%' OR [SheetName] LIKE '%{0}%' OR "
            u"[Revision] LIKE '%{0}%' OR [NuevaRevision] LIKE '%{0}%'"
        ).format(esc)
    _sigrev_sync_sel_all_hdr(state)


def _sigrev_on_buscar_changed(state, sender, args):
    _sigrev_sync_buscar_wm(state)
    _sigrev_apply_buscar_filter(state)


def _sigrev_bind_sheet_table(state, tbl):
    _sigrev_detach_row_changed(state)
    state["sheet_table"] = tbl
    state["_sigrev_row_delegate"] = _sigrev_make_row_delegate(state)
    tbl.RowChanged += state["_sigrev_row_delegate"]


def _sigrev_on_table_row_changed(state, sender, args):
    if state.get("_sigrev_syncing_sel_all"):
        return
    _sigrev_refresh_estado_sheet(state)
    _sigrev_sync_sel_all_hdr(state)


def _sigrev_find_header_select_all_checkbox(original_source):
    """Localiza el CheckBox de «marcar todas» en la cabecera del DataGrid (Tag HdrSelectAll)."""
    try:
        from System.Windows.Controls import CheckBox
        from System.Windows.Media import VisualTreeHelper
    except Exception:
        return None
    d = original_source
    for _ in range(48):
        if d is None:
            break
        try:
            if isinstance(d, CheckBox):
                tg = d.Tag
                if tg is not None and unicode(str(tg)).strip() == u"HdrSelectAll":
                    return d
        except Exception:
            pass
        try:
            d = VisualTreeHelper.GetParent(d)
        except Exception:
            break
    return None


def _sigrev_apply_select_all_visible_rows(state):
    """
    Activa o desactiva Sel en todas las filas visibles (DefaultView / filtro Buscar).
    Usa la misma ruta que el clic por fila (``DataGrid.Items[i][Sel]``) para que el binding actualice la UI.
    """
    tbl = state.get("sheet_table")
    g = state.get("_sigrev_grid")
    if tbl is None:
        return
    dv = tbl.DefaultView
    try:
        n = int(dv.Count)
    except Exception:
        return
    try:
        from System import Boolean as ClrBool
    except Exception:
        return
    if n == 0:
        _sigrev_refresh_estado_sheet(state)
        _sigrev_sync_sel_all_hdr(state)
        return
    n_eligible = 0
    n_sel_among_eligible = 0
    for i in range(n):
        try:
            rv = g.Items[i] if g is not None else dv[i]
            if not _sigrev_row_sel_enabled_rowview(rv):
                continue
            n_eligible += 1
            if _sigrev_row_selected_rowview(rv):
                n_sel_among_eligible += 1
        except Exception:
            pass
    new_val = True
    if n_eligible > 0 and n_sel_among_eligible == n_eligible:
        new_val = False
    state["_sigrev_syncing_sel_all"] = True
    try:
        for i in range(n):
            try:
                rv = g.Items[i] if g is not None else dv[i]
                if not _sigrev_row_sel_enabled_rowview(rv):
                    continue
                if g is not None:
                    g.Items[i][u"Sel"] = ClrBool(new_val)
                else:
                    dv[i][u"Sel"] = ClrBool(new_val)
            except Exception:
                pass
    finally:
        state["_sigrev_syncing_sel_all"] = False
    _sigrev_refresh_estado_sheet(state)
    _sigrev_sync_sel_all_hdr(state)


def _sigrev_wire_select_all_hdr(state):
    try:
        from System.Windows.Media import VisualTreeHelper
        from System.Windows.Controls import CheckBox
    except Exception:
        return
    grid = state.get("_sigrev_grid")
    if grid is None:
        return

    def walk_find(comp):
        try:
            n = VisualTreeHelper.GetChildrenCount(comp)
        except Exception:
            return None
        for i in range(int(n)):
            try:
                ch = VisualTreeHelper.GetChild(comp, i)
            except Exception:
                continue
            try:
                if isinstance(ch, CheckBox):
                    tg = ch.Tag
                    if tg is not None and unicode(str(tg)) == u"HdrSelectAll":
                        return ch
            except Exception:
                pass
            got = walk_find(ch)
            if got is not None:
                return got
        return None

    chk = walk_find(grid)
    if chk is None:
        return
    state["_sigrev_hdr_chk"] = chk
    # Marcar todas: la lógica va en PreviewMouseLeftButtonDown del DataGrid
    # (_sigrev_on_grid_preview_mouse) para que los datos Sel y el checkbox del encabezado no se desincronicen.


def _sigrev_sync_sel_all_hdr(state):
    chk = state.get("_sigrev_hdr_chk")
    if chk is None:
        try:
            _sigrev_wire_select_all_hdr(state)
        except Exception:
            pass
        chk = state.get("_sigrev_hdr_chk")
    if chk is None:
        return
    tbl = state.get("sheet_table")
    if tbl is None:
        return
    dv = tbl.DefaultView
    try:
        n = int(dv.Count)
    except Exception:
        n = 0
    state["_sigrev_syncing_sel_all"] = True
    try:
        try:
            chk.IsThreeState = True
        except Exception:
            pass
        if n == 0:
            chk.IsChecked = False
            return
        n_eligible = 0
        n_sel_among_eligible = 0
        for i in range(n):
            try:
                rv = dv[i]
                if not _sigrev_row_sel_enabled_rowview(rv):
                    continue
                n_eligible += 1
                if _sigrev_row_selected_rowview(rv):
                    n_sel_among_eligible += 1
            except Exception:
                pass
        if n_eligible == 0:
            chk.IsChecked = False
            return
        if n_sel_among_eligible == 0:
            chk.IsChecked = False
        elif n_sel_among_eligible == n_eligible:
            chk.IsChecked = True
        else:
            chk.IsChecked = None
    finally:
        state["_sigrev_syncing_sel_all"] = False


def _sigrev_on_grid_loaded_sheet(state, sender, args):
    from System.Windows import Visibility

    g = sender
    try:
        if g.Columns.Count > 0:
            g.Columns[g.Columns.Count - 1].Visibility = Visibility.Collapsed
    except Exception:
        pass
    try:
        from System import Action
        from System.Windows.Threading import DispatcherPriority

        def _wire():
            try:
                _sigrev_wire_select_all_hdr(state)
                _sigrev_sync_sel_all_hdr(state)
            except Exception:
                pass

        _wire()
        win = state.get("win")
        if win is not None:
            g.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(_wire))
            g.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_wire))
    except Exception:
        try:
            _sigrev_wire_select_all_hdr(state)
            _sigrev_sync_sel_all_hdr(state)
        except Exception:
            pass
    try:
        win = state["win"]
        _sigrev_schedule_sheet_grid_scrollbars(win, g)
    except Exception:
        pass


def _sigrev_on_grid_preview_mouse(state, sender, e):
    try:
        from System.Windows.Controls import DataGridCell, DataGridCheckBoxColumn
        from System.Windows.Input import Keyboard, ModifierKeys
        from System.Windows.Media import VisualTreeHelper
        from System import Boolean as ClrBool
    except Exception:
        return
    tbl = state.get("sheet_table")
    g = state.get("_sigrev_grid")
    if tbl is None or g is None:
        return

    if _sigrev_find_header_select_all_checkbox(e.OriginalSource) is not None:
        if not state.get("_sigrev_syncing_sel_all"):
            _sigrev_apply_select_all_visible_rows(state)
        try:
            e.Handled = True
        except Exception:
            pass
        return

    d = e.OriginalSource
    cell = None
    for _ in range(40):
        if d is None:
            break
        try:
            if isinstance(d, DataGridCell):
                cell = d
                break
        except Exception:
            pass
        try:
            d = VisualTreeHelper.GetParent(d)
        except Exception:
            break
    if cell is None:
        return
    try:
        if not isinstance(cell.Column, DataGridCheckBoxColumn):
            return
    except Exception:
        return
    drv = cell.DataContext
    if drv is None:
        return
    if not _sigrev_row_sel_enabled_rowview(drv):
        try:
            e.Handled = True
        except Exception:
            pass
        return
    idx = _sigrev_grid_index_of_datarowview(g, drv)
    if idx < 0:
        return

    def _sel_get(rowv):
        return _sigrev_row_selected_rowview(rowv)

    shift_down = False
    try:
        shift_down = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
    except Exception:
        pass
    anchor = state.get("_sigrev_sel_anchor")

    if shift_down and anchor is not None:
        try:
            cur_on = _sel_get(drv)
            new_val = not cur_on
            i0 = min(int(anchor), idx)
            i1 = max(int(anchor), idx)
            for j in range(i0, i1 + 1):
                try:
                    rvj = g.Items[j]
                    if not _sigrev_row_sel_enabled_rowview(rvj):
                        continue
                    g.Items[j][u"Sel"] = ClrBool(new_val)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            e.Handled = True
        except Exception:
            pass
        _sigrev_refresh_estado_sheet(state)
        _sigrev_sync_sel_all_hdr(state)
        return

    def _bulk():
        try:
            cur_on = _sel_get(drv)
            new_val = not cur_on
            state["_sigrev_sel_anchor"] = idx
            indices = []
            seen = set()
            try:
                coll = g.SelectedItems
                if coll is not None:
                    for i in range(int(coll.Count)):
                        try:
                            it = coll[i]
                            j = _sigrev_grid_index_of_datarowview(g, it)
                            if j >= 0 and j not in seen:
                                seen.add(j)
                                indices.append(j)
                        except Exception:
                            pass
            except Exception:
                pass
            if idx not in seen:
                indices.append(idx)
            for j in indices:
                try:
                    rvj = g.Items[j]
                    if not _sigrev_row_sel_enabled_rowview(rvj):
                        continue
                    g.Items[j][u"Sel"] = ClrBool(new_val)
                except Exception:
                    pass
        except Exception:
            pass
        _sigrev_refresh_estado_sheet(state)
        _sigrev_sync_sel_all_hdr(state)

    try:
        from System import Action
        from System.Windows.Threading import DispatcherPriority

        g.Dispatcher.BeginInvoke(DispatcherPriority.Input, Action(_bulk))
    except Exception:
        _bulk()
    try:
        e.Handled = True
    except Exception:
        pass


def _sigrev_rebind_sheets_grid(state, reset_buscar=False):
    doc = state["doc"]
    win = state["win"]
    sheets_all = _collect_sheets(doc)
    tbl = _build_sheet_selection_table(doc, sheets_all)
    _sigrev_bind_sheet_table(state, tbl)
    grid = state["_sigrev_grid"]
    grid.ItemsSource = tbl.DefaultView
    if reset_buscar:
        try:
            win.FindName("TxtBuscar").Text = u""
        except Exception:
            pass
    _sigrev_apply_buscar_filter(state)
    _sigrev_sync_buscar_wm(state)
    _sigrev_refresh_estado_sheet(state)
    _sigrev_refresh_nueva_revision_from_ui(state)


def _dates_dd_mm_yy_options():
    """
    Fechas en formato dd.MM.yy: desde 5 días antes de hoy hasta 20 días después (inclusive).

    Returns:
        (lista de cadenas, indice_del_dia_actual) para usar como SelectedIndex del combo.
    """
    dias_antes = 5
    dias_despues = 20
    today = DateTime.Today
    fmt = "dd.MM.yy"
    inv = CultureInfo.InvariantCulture
    opts = []
    for i in range(-dias_antes, dias_despues + 1):
        d = today.AddDays(i)
        opts.append(d.ToString(fmt, inv))
    return opts, dias_antes


def _parse_dd_mm_yy(s):
    return DateTime.ParseExact(
        unicode(s).strip(), "dd.MM.yy", CultureInfo.InvariantCulture
    )


def _revisions_on_sheet_ordered(doc, sheet):
    revs = []
    for eid in sheet.GetAdditionalRevisionIds():
        el = doc.GetElement(eid)
        if isinstance(el, Revision):
            revs.append(el)
    return revs


def _element_id_integer(eid):
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return 0


def _revision_ids_ordered_in_project(doc):
    """Orden igual que Gestión de revisiones (API: Revision.GetAllRevisionIds)."""
    try:
        raw = Revision.GetAllRevisionIds(doc)
    except Exception:
        raw = None
    out = []
    if raw:
        for eid in raw:
            if eid is not None and eid != ElementId.InvalidElementId:
                out.append(eid)
    return out


def _index_revision_in_project_order(project_ordered_ids, revision_element):
    t = _element_id_integer(revision_element.Id)
    for i, eid in enumerate(project_ordered_ids):
        if _element_id_integer(eid) == t:
            return i
    return -1


def _furthest_sheet_revision_on_project_sequence(existing_revs, project_ordered_ids):
    """
    Entre todas las revisiones colocadas en la lámina, la que aparece más adelante
    en GetAllRevisionIds (equivale a la última revisión efectiva en el proyecto).
    """
    best = None
    best_i = -1
    for r in existing_revs:
        j = _index_revision_in_project_order(project_ordered_ids, r)
        if j < 0:
            continue
        if j > best_i:
            best_i = j
            best = r
    return best, best_i



def _revision_number_display(rev_el):
    """Número/letra de la fila de revisión (p. ej. 0, 1, A) — alineado con listado_planos_excel_core."""
    try:
        s = unicode(getattr(rev_el, u"RevisionNumber", None) or u"").strip()
        if s:
            return s
    except Exception:
        pass
    for nm in ("REVIT_REVISION_NUMBER", "REVISION_NUMBER"):
        bip = getattr(BuiltInParameter, nm, None)
        if bip is None:
            continue
        try:
            p = rev_el.get_Parameter(bip)
        except Exception:
            p = None
        if p is None:
            continue
        try:
            s = (p.AsString() or p.AsValueString() or u"").strip()
        except Exception:
            s = u""
        if s:
            return s
    return u""


def _index_revision_project_display_number(doc, ordered_ids, want_display):
    """Índice en ``Revision.GetAllRevisionIds`` de la primera revisión cuyo número mostrado coincide (p. ej. «0»)."""
    w = unicode(want_display or u"").strip()
    if not w:
        return -1
    wl = w.lower()
    for i, eid in enumerate(ordered_ids):
        rev = doc.GetElement(eid)
        if not isinstance(rev, Revision):
            continue
        num = (_revision_number_display(rev) or u"").strip()
        if not num:
            continue
        if num == w or num.lower() == wl:
            return int(i)
        try:
            if int(num) == int(w):
                return int(i)
        except Exception:
            pass
    return -1


def _sigrev_compute_sel_enabled_for_sheet(doc, sheet, ordered, revision_emit_revision_zero, ti_rev0):
    """False solo en modo Revisión 0 cuando la última revisión del índice ya es la número 0 del proyecto."""
    if not revision_emit_revision_zero:
        return True
    if ti_rev0 < 0 or not ordered:
        return True
    existing = _revisions_on_sheet_ordered(doc, sheet)
    if existing:
        _furthest, fi = _furthest_sheet_revision_on_project_sequence(existing, ordered)
        if _furthest is None or fi < 0:
            return True
    else:
        fi = -1
    return not (fi >= 0 and fi == ti_rev0)


def _sigrev_row_sel_enabled_rowview(rv):
    try:
        if rv is None:
            return True
        if hasattr(rv, "Row"):
            try:
                return bool(rv.Row[u"SelEnabled"])
            except Exception:
                pass
        try:
            return bool(rv[u"SelEnabled"])
        except Exception:
            pass
    except Exception:
        pass
    return True


def _sigrev_nueva_revision_preview_for_sheet(doc, sheet, ordered, revision_emit_revision_zero, ti_rev0):
    """Número mostrado de la revisión que se aplicaría (misma lógica que la emisión)."""
    if not ordered:
        return u"—"
    existing = _revisions_on_sheet_ordered(doc, sheet)
    if existing:
        furthest, fi = _furthest_sheet_revision_on_project_sequence(existing, ordered)
        if furthest is None or fi < 0:
            return u"(Sin coincidencia)"
    else:
        fi = -1
    ti = fi + 1
    if revision_emit_revision_zero:
        if ti_rev0 < 0:
            return u"(Sin rev. 0)"
        if fi >= 0 and fi == ti_rev0:
            return u"(Ya en rev. 0)"
        ti = ti_rev0
    if ti >= len(ordered):
        return u"(Sin siguiente)"
    tgt = doc.GetElement(ordered[ti])
    if not isinstance(tgt, Revision):
        return u"—"
    s = (_revision_number_display(tgt) or u"").strip()
    return s if s else u"—"


def _sheet_r_slot_prefix(slot_1based):
    """R09_ para slot_1based=9."""
    return u"R{:02d}_".format(int(slot_1based))


def _title_block_instances_on_sheet(doc, sheet):
    """
    Cajetines colocados en la lámina. Los parámetros Rnn_* suelen estar aquí, no en ViewSheet.
    """
    bic = getattr(BuiltInCategory, "OST_TitleBlocks", None)
    if bic is None or sheet is None or doc is None:
        return []
    try:
        return list(
            FilteredElementCollector(doc, sheet.Id)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []


def _lookup_parameter_sheet_then_titleblocks(sheet, doc, param_name):
    """
    Busca el parámetro primero en la lámina y después en cada instancia de cajetín.

    Devuelve (elemento_host, Parameter) o (None, None).
    """
    if not param_name or sheet is None or doc is None:
        return None, None
    try:
        p = sheet.LookupParameter(param_name)
        if p is not None:
            return sheet, p
    except Exception:
        pass
    for tb in _title_block_instances_on_sheet(doc, sheet):
        try:
            p = tb.LookupParameter(param_name)
            if p is not None:
                return tb, p
        except Exception:
            continue
    return None, None


def _detect_revision_param_layout(sheet, doc):
    """
    Dos convenciones de nombres en cajetín:

    - ``rnn_field``: ``R02_01_NUM``, ``R02_02_DES``, … (fila en prefijo ``Rnn``).
    - ``r01_row``: ``R01_02_NUM``, ``R01_02_DES``, … (todas las filas bajo ``R01``).
    """
    if sheet is None or doc is None:
        return u"rnn_field"
    _, p_r2 = _lookup_parameter_sheet_then_titleblocks(sheet, doc, u"R02_01_NUM")
    if p_r2 is not None:
        return u"rnn_field"
    _, p_12 = _lookup_parameter_sheet_then_titleblocks(sheet, doc, u"R01_02_NUM")
    if p_12 is not None:
        return u"r01_row"
    _, p_11 = _lookup_parameter_sheet_then_titleblocks(sheet, doc, u"R01_01_NUM")
    if p_11 is not None:
        return u"r01_row"
    return u"rnn_field"


def _revision_num_parameter_name(layout, slot_1based):
    s = int(slot_1based)
    if layout == u"r01_row":
        return u"R01_{:02d}_NUM".format(s)
    return _sheet_r_slot_prefix(s) + RNN_SUFFIX_NUM


def _revision_slot_display(layout, slot_1based):
    """Etiqueta para mensajes (p. ej. R01_03 vs R03)."""
    s = int(slot_1based)
    if layout == u"r01_row":
        return u"R01_{:02d}".format(s)
    return u"R{:02d}".format(s)


def _slot_1based_in_sheet_issue_list(ids_list, next_revision_id_int):
    """
    Posición 1-based del ElementId de revisión en la lista de índice de lámina
    (orden de ``SetAdditionalRevisionIds``): R01 = primera emisión de esa lámina.
    """
    try:
        n = int(ids_list.Count)
    except Exception:
        return 0
    want = int(next_revision_id_int)
    for i in range(n):
        try:
            eid = ids_list[i]
        except Exception:
            try:
                eid = ids_list.get_Item(i)
            except Exception:
                continue
        if _element_id_integer(eid) == want:
            return i + 1
    return 0


def _rnn_num_parameter_text_trimmed(sheet, slot_1based, doc, layout):
    """
    Valor textual del parámetro NUM de la fila ``slot_1based``; None si no existe.
    """
    from Autodesk.Revit.DB import StorageType

    _, p = _lookup_parameter_sheet_then_titleblocks(
        sheet,
        doc,
        _revision_num_parameter_name(layout, slot_1based),
    )
    if p is None:
        return None
    try:
        if p.StorageType == StorageType.String:
            return (unicode(p.AsString()) or u"").strip()
        return (unicode(p.AsValueString() or u"")).strip()
    except Exception:
        try:
            return (unicode(p.AsValueString() or u"")).strip()
        except Exception:
            return u""


def _revision_num_slot_empty_for_emit(sheet, doc, slot_1based, layout):
    """
    True si el NUM de esa fila existe y puede recibir una nueva emisión.
    """
    _, p = _lookup_parameter_sheet_then_titleblocks(
        sheet,
        doc,
        _revision_num_parameter_name(layout, slot_1based),
    )
    if p is None:
        return False
    try:
        if not p.HasValue:
            return True
    except Exception:
        pass
    txt = _rnn_num_parameter_text_trimmed(sheet, slot_1based, doc, layout)
    if txt is None:
        return False
    return not txt


def _first_empty_rnn_num_slot_through(sheet, through_slot_1based, doc, layout):
    """Primera fila en [1 … through_slot] con NUM vacío; −1 si ninguna."""
    last = max(1, min(int(through_slot_1based), MAX_REVISION_SLOTS))
    for s in range(1, last + 1):
        if _revision_num_slot_empty_for_emit(sheet, doc, s, layout):
            return s
    return -1


def _parameter_set_stringish(p, value):
    """Asigna texto según StorageType del parámetro Revit."""
    from Autodesk.Revit.DB import StorageType

    try:
        uval = unicode(value) if value is not None else u""
    except Exception:
        uval = u""
    if p is None or p.IsReadOnly:
        return False
    try:
        if p.StorageType == StorageType.String:
            p.Set(uval)
            return True
        if p.StorageType == StorageType.Integer:
            try:
                p.Set(int(round(float(uval.strip() or 0))))
            except Exception:
                return False
            return True
        if p.StorageType == StorageType.Double:
            try:
                p.Set(float(uval.strip().replace(u",", u".")))
            except Exception:
                return False
            return True
    except Exception:
        return False
    return False


def _try_set_sheet_or_tb_parameter(sheet, doc, full_name, value):
    """Establece parámetro por nombre en lámina o cajetín (donde exista)."""
    _, p = _lookup_parameter_sheet_then_titleblocks(sheet, doc, full_name)
    if p is None:
        return False
    return _parameter_set_stringish(p, value)


def _try_set_slot_dibujo(sheet, doc, layout, slot_1based, value):
    """Campo dibujó según layout (``03_DIR``/``03_DIB`` o ``R01_mm_DIR``/``DIB``)."""
    row = int(slot_1based)
    if layout == u"r01_row":
        ok_dir = _try_set_sheet_or_tb_parameter(
            sheet, doc, u"R01_{:02d}_DIR".format(row), value
        )
        ok_dib = _try_set_sheet_or_tb_parameter(
            sheet, doc, u"R01_{:02d}_DIB".format(row), value
        )
        return ok_dir or ok_dib
    pref = _sheet_r_slot_prefix(row)
    ok_dir = _try_set_sheet_or_tb_parameter(sheet, doc, pref + RNN_SUFFIX_DIR, value)
    ok_dib = _try_set_sheet_or_tb_parameter(sheet, doc, pref + RNN_SUFFIX_DIB, value)
    return ok_dir or ok_dib


def _set_sheet_revision_rnn_slot(
    sheet,
    doc,
    layout,
    slot_1based,
    numero_str,
    descripcion,
    dibujo,
    reviso,
    aprobo,
    fecha_dd_mm_yy,
):
    """
    Escribe los seis campos de una fila de revisión en lámina/cajetín.

    Soporta convención ``Rnn_01_NUM``… y ``R01_mm_NUM``… (``layout``).
    """
    row = int(slot_1based)
    dib_eff = (dibujo or u"").strip()
    if not dib_eff:
        dib_eff = (reviso or u"").strip()
    if layout == u"r01_row":
        _try_set_sheet_or_tb_parameter(
            sheet, doc, u"R01_{:02d}_NUM".format(row), numero_str or u""
        )
        _try_set_sheet_or_tb_parameter(
            sheet, doc, u"R01_{:02d}_DES".format(row), descripcion or u""
        )
        _try_set_slot_dibujo(sheet, doc, layout, row, dib_eff)
        _try_set_sheet_or_tb_parameter(
            sheet, doc, u"R01_{:02d}_REV".format(row), reviso or u""
        )
        _try_set_sheet_or_tb_parameter(
            sheet, doc, u"R01_{:02d}_APR".format(row), aprobo or u""
        )
        _try_set_sheet_or_tb_parameter(
            sheet, doc, u"R01_{:02d}_FCH".format(row), fecha_dd_mm_yy or u""
        )
        return
    pref = _sheet_r_slot_prefix(row)
    _try_set_sheet_or_tb_parameter(sheet, doc, pref + RNN_SUFFIX_NUM, numero_str or u"")
    _try_set_sheet_or_tb_parameter(sheet, doc, pref + RNN_SUFFIX_DES, descripcion or u"")
    _try_set_slot_dibujo(sheet, doc, layout, row, dib_eff)
    _try_set_sheet_or_tb_parameter(sheet, doc, pref + RNN_SUFFIX_REV, reviso or u"")
    _try_set_sheet_or_tb_parameter(sheet, doc, pref + RNN_SUFFIX_APR, aprobo or u"")
    _try_set_sheet_or_tb_parameter(sheet, doc, pref + RNN_SUFFIX_FCH, fecha_dd_mm_yy or u"")


def _iter_revision_clouds_on_sheet(doc, sheet_id):
    """
    BuiltInCategory debe ser OST_RevisionClouds (con guión bajo tras OST).

    En hosts viejos si falla el enum, se filtra por clase RevisionCloud.
    """
    bic = getattr(BuiltInCategory, "OST_RevisionClouds", None)
    if bic is not None:
        try:
            for c in (
                FilteredElementCollector(doc, sheet_id)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
            ):
                yield c
            return
        except Exception:
            pass
    if RevisionCloud is None:
        return
    try:
        for c in (
            FilteredElementCollector(doc, sheet_id)
            .OfClass(RevisionCloud)
            .WhereElementIsNotElementType()
        ):
            yield c
    except Exception:
        return


def _update_cantidad_revisiones(doc, sheet, count):
    from Autodesk.Revit.DB import StorageType

    for c in _iter_revision_clouds_on_sheet(doc, sheet.Id):
        p = c.LookupParameter(PARAM_CANTIDAD_REVISIONES)
        if p is None or p.IsReadOnly:
            continue
        try:
            if p.StorageType == StorageType.String:
                p.Set(unicode(count))
            elif p.StorageType == StorageType.Integer:
                p.Set(int(count))
        except Exception:
            pass


def _apply_revision_to_sheets(
    doc,
    sheets,
    description,
    dibujo,
    reviso,
    aprobo,
    fecha_dd_mm_yy_str,
    revision_emit_revision_zero=False,
    revit_uiapp=None,
):
    done = 0
    errs = []
    ordered = _revision_ids_ordered_in_project(doc)
    ti_rev0 = -1
    if revision_emit_revision_zero:
        ti_rev0 = _index_revision_project_display_number(doc, ordered, u"0")
        if ti_rev0 < 0:
            return (
                0,
                u"No existe ninguna revisión con número 0 en Gestión de revisiones del proyecto.",
            )
    sheets_list = list(sheets)
    ntot = len(sheets_list)
    pb = None
    pb_ok = False
    blocker = None
    blocker_ok = False
    try:
        if revit_uiapp is not None and ntot > 0:
            if _BloquearComandosRevit is not None:
                try:
                    blocker = _BloquearComandosRevit(revit_uiapp)
                    blocker.__enter__()
                    blocker_ok = True
                except Exception:
                    blocker = None
                    blocker_ok = False
            try:
                pb = _sigrev_pbar_start(
                    _sigrev_pbar_initial_title(_SIGREV_PBAR_TITLE_BASE, ntot),
                    ntot,
                )
                if pb is not None:
                    pb.__enter__()
                    pb_ok = True
            except Exception:
                pb = None
                pb_ok = False

        tx_mega = None
        pending_writes = 0
        aborted_tx = False

        for si, sheet in enumerate(sheets_list):
            try:
                if not ordered:
                    errs.append(
                        u"{}: el proyecto no define revisiones (Gestión de revisiones)."
                        .format(_sheet_display(sheet))
                    )
                    continue

                existing = _revisions_on_sheet_ordered(doc, sheet)
                if existing:
                    furthest, fi = _furthest_sheet_revision_on_project_sequence(existing, ordered)
                    if furthest is None or fi < 0:
                        errs.append(
                            u"{}: las revisiones de la lámina no coinciden con la secuencia del proyecto.".format(
                                _sheet_display(sheet)
                            )
                        )
                        continue
                else:
                    furthest = None
                    fi = -1

                if revision_emit_revision_zero:
                    if fi >= 0 and fi == ti_rev0:
                        errs.append(
                            u"{}: modo revisión 0 omitido: la última revisión del índice de la lámina ya es la revisión 0 del proyecto.".format(
                                _sheet_display(sheet)
                            )
                        )
                        continue

                ti = fi + 1
                if revision_emit_revision_zero:
                    ti = ti_rev0

                if ti >= len(ordered):
                    errs.append(
                        u"{}: no hay revisión válida después de la última del índice. Defínala en Gestión de revisiones.".format(
                            _sheet_display(sheet)
                        )
                    )
                    continue

                target_rev_id = ordered[ti]

                if tx_mega is None:
                    tx_mega = Transaction(doc, u"Arainco: Revisiones")
                    tx_mega.Start()

                target_rev = doc.GetElement(target_rev_id)
                if not isinstance(target_rev, Revision):
                    raise Exception(u"ElementId objetivo no es una revisión.")

                ids = ClrList[ElementId]()
                for eid in sheet.GetAdditionalRevisionIds():
                    ids.Add(eid)

                on_sheet_ids = set()
                nc_ids = int(ids.Count)
                for i in range(nc_ids):
                    try:
                        xe = ids[i]
                    except Exception:
                        try:
                            xe = ids.get_Item(i)
                        except Exception:
                            continue
                    on_sheet_ids.add(_element_id_integer(xe))

                for j in range(fi + 1, ti + 1):
                    ej = ordered[j]
                    jid = _element_id_integer(ej)
                    if jid not in on_sheet_ids:
                        ids.Add(ej)
                        on_sheet_ids.add(jid)

                ni_target = _element_id_integer(target_rev_id)
                geom_slot = _slot_1based_in_sheet_issue_list(ids, ni_target)
                if geom_slot < 1:
                    raise Exception(u"No se pudo ubicar la revisión en la lista del índice de lámina.")

                layout = _detect_revision_param_layout(sheet, doc)
                slot_write_rev0 = None
                slot_auto = None

                if revision_emit_revision_zero:
                    slot_write_rev0 = _first_empty_rnn_num_slot_through(
                        sheet, MAX_REVISION_SLOTS, doc, layout
                    )
                    if slot_write_rev0 < 1:
                        raise Exception(
                            u"No hay ninguna fila R01–R{} con NUM vacío en esta lámina/cajetín.".format(
                                MAX_REVISION_SLOTS
                            )
                        )
                    _, p_chk = _lookup_parameter_sheet_then_titleblocks(
                        sheet,
                        doc,
                        _revision_num_parameter_name(layout, slot_write_rev0),
                    )
                    if p_chk is None:
                        raise Exception(
                            u"No existe el parámetro {} en lámina ni cajetín (¿plantilla?).".format(
                                _revision_num_parameter_name(layout, slot_write_rev0)
                            )
                        )
                else:
                    slot_auto = _first_empty_rnn_num_slot_through(
                        sheet, MAX_REVISION_SLOTS, doc, layout
                    )
                    if slot_auto < 1:
                        raise Exception(
                            u"No hay ninguna fila R01–R{} con NUM vacío en esta lámina/cajetín.".format(
                                MAX_REVISION_SLOTS
                            )
                        )
                    _, p_auto = _lookup_parameter_sheet_then_titleblocks(
                        sheet,
                        doc,
                        _revision_num_parameter_name(layout, slot_auto),
                    )
                    if p_auto is None:
                        raise Exception(
                            u"No existe el parámetro {} en lámina ni cajetín (¿plantilla?).".format(
                                _revision_num_parameter_name(layout, slot_auto)
                            )
                        )

                sheet.SetAdditionalRevisionIds(ids)

                if revision_emit_revision_zero:
                    numero = _revision_number_display(target_rev)
                    _set_sheet_revision_rnn_slot(
                        sheet,
                        doc,
                        layout,
                        int(slot_write_rev0),
                        numero,
                        description,
                        dibujo,
                        reviso,
                        aprobo,
                        fecha_dd_mm_yy_str,
                    )
                    if geom_slot != slot_write_rev0:
                        errs.append(
                            u"{}: revisión 0 en la posición {} del índice de Revit; "
                            u"los datos del formulario quedaron en {} (primera fila libre del cajetín).".format(
                                _sheet_display(sheet),
                                geom_slot,
                                _revision_slot_display(layout, slot_write_rev0),
                            )
                        )
                else:
                    numero_final = _revision_number_display(target_rev)
                    _set_sheet_revision_rnn_slot(
                        sheet,
                        doc,
                        layout,
                        int(slot_auto),
                        numero_final,
                        description,
                        dibujo,
                        reviso,
                        aprobo,
                        fecha_dd_mm_yy_str,
                    )
                    if geom_slot != slot_auto:
                        errs.append(
                            u"{}: la nueva revisión está en la posición {} del índice de Revit; "
                            u"los datos del formulario quedaron en {} (primera fila libre del cajetín).".format(
                                _sheet_display(sheet),
                                geom_slot,
                                _revision_slot_display(layout, slot_auto),
                            )
                        )
                _update_cantidad_revisiones(doc, sheet, ids.Count)
                pending_writes += 1
            except Exception as ex:
                pending_writes = 0
                aborted_tx = True
                errs.append(u"{}: {}".format(_sheet_display(sheet), unicode(ex)))
                try:
                    if (
                        tx_mega is not None
                        and tx_mega.HasStarted()
                        and not tx_mega.HasEnded()
                    ):
                        tx_mega.RollBack()
                except Exception:
                    pass
                break
            finally:
                if pb_ok and pb is not None:
                    try:
                        _sigrev_pbar_step(pb, si, ntot, _SIGREV_PBAR_TITLE_BASE)
                    except Exception:
                        pass

        if (
            not aborted_tx
            and tx_mega is not None
            and tx_mega.HasStarted()
            and not tx_mega.HasEnded()
        ):
            try:
                tx_mega.Commit()
                done = pending_writes
            except Exception as ex_commit:
                errs.append(
                    u"No se pudo confirmar la transacción agrupada: {0}".format(
                        unicode(ex_commit)
                    )
                )
                try:
                    if tx_mega.HasStarted() and not tx_mega.HasEnded():
                        tx_mega.RollBack()
                except Exception:
                    pass
    finally:
        _sigrev_pb_exit_safe(pb, pb_ok)
        if blocker_ok and blocker is not None:
            try:
                blocker.__exit__(None, None, None)
            except Exception:
                pass

    return done, u"\n".join(errs)


def _try_activate_existing():
    """
    Instancia única: solo bloquea si la ventana registrada sigue visible.
    Referencias débiles pueden dejar objetos vivos hasta GC; evitamos ese falso positivo.
    """
    try:
        o = AppDomain.CurrentDomain.GetData(APP_DOMAIN_SINGLETON_KEY)
        if o is None:
            return False
        win = o.Target if hasattr(o, "Target") else o
        if win is None:
            AppDomain.CurrentDomain.SetData(APP_DOMAIN_SINGLETON_KEY, None)
            return False
        try:
            if not getattr(win, "IsVisible", False):
                AppDomain.CurrentDomain.SetData(APP_DOMAIN_SINGLETON_KEY, None)
                return False
            win.WindowState = WindowState.Normal
            win.Activate()
            return True
        except Exception:
            AppDomain.CurrentDomain.SetData(APP_DOMAIN_SINGLETON_KEY, None)
            return False
    except Exception:
        return False


def _singleton_register(win):
    AppDomain.CurrentDomain.SetData(APP_DOMAIN_SINGLETON_KEY, win)


def _singleton_clear():
    try:
        AppDomain.CurrentDomain.SetData(APP_DOMAIN_SINGLETON_KEY, None)
    except Exception:
        pass


def _sigrev_close_with_fade(win):
    """
    Fade-out + slide-down sobre SigRevAnimShell; luego Close().
    Sin Window.Opacity animado ni evento Closing (evita bloqueos ShowDialog + IronPython).
    Completed + DispatcherTimer garantizan Close().
    """
    global _SIGREV_CLOSE_GUARD
    if _SIGREV_CLOSE_GUARD.get("busy"):
        return
    _SIGREV_CLOSE_GUARD["busy"] = True
    _SIGREV_CLOSE_GUARD["finalized"] = False

    def _cleanup_anim_targets():
        try:
            from System.Windows import UIElement
            from System.Windows.Media import TranslateTransform

            try:
                win.BeginAnimation(UIElement.OpacityProperty, None)
            except Exception:
                pass
            ch_clr = win.FindName("SigRevRootChrome")
            if ch_clr is not None:
                try:
                    ch_clr.BeginAnimation(UIElement.OpacityProperty, None)
                except Exception:
                    pass
            shell = win.FindName("SigRevAnimShell")
            if shell is not None:
                try:
                    shell.BeginAnimation(UIElement.OpacityProperty, None)
                except Exception:
                    pass
            tt = win.FindName("SigRevEnterTranslate")
            if tt is not None:
                try:
                    tt.BeginAnimation(TranslateTransform.YProperty, None)
                except Exception:
                    pass
        except Exception:
            pass

    def _do_close():
        if _SIGREV_CLOSE_GUARD.get("finalized"):
            return
        _SIGREV_CLOSE_GUARD["finalized"] = True
        _cleanup_anim_targets()
        try:
            win.Close()
        except Exception:
            pass
        finally:
            try:
                _SIGREV_CLOSE_GUARD["busy"] = False
            except Exception:
                pass

    done = {"ok": False}

    def _finish_once(_s=None, _e=None):
        if done["ok"]:
            return
        done["ok"] = True
        try:
            tm = getattr(win, "_sigrev_exit_fallback_timer", None)
            if tm is not None:
                try:
                    tm.Stop()
                except Exception:
                    pass
        except Exception:
            pass
        _do_close()

    try:
        from System import TimeSpan
        from System.Windows.Threading import DispatcherTimer

        _sigrev_snap_anim_shell_to_visible(win)

        sb_ex = win.TryFindResource("SigRevExitStoryboard")
        shell = win.FindName("SigRevAnimShell")
        tt = win.FindName("SigRevEnterTranslate")
        if sb_ex is None or shell is None or tt is None:
            _finish_once()
            return

        try:
            from System.Windows import UIElement
            from System.Windows.Media import TranslateTransform

            shell.BeginAnimation(UIElement.OpacityProperty, None)
            tt.BeginAnimation(TranslateTransform.YProperty, None)
        except Exception:
            pass

        dur = Duration(TimeSpan.FromMilliseconds(float(_SIGREV_CHROME_MS)))
        try:
            for i in range(int(sb_ex.Children.Count)):
                sb_ex.Children[i].Duration = dur
        except Exception:
            pass

        sb_ex.Completed += EventHandler(_finish_once)
        sb_ex.Begin(win, True)

        tm = DispatcherTimer()
        tm.Interval = TimeSpan.FromMilliseconds(float(_SIGREV_EXIT_FALLBACK_MS))
        tm.Tick += EventHandler(lambda _snd, _evt: _finish_once())
        win._sigrev_exit_fallback_timer = tm
        tm.Start()
    except Exception:
        _finish_once()


class _RevWindowHandlers(object):
    """Code-behind WPF routed handlers (IronPython retains weak refs to bound methods)."""

    def __init__(self, state):
        self._s = state

    def cancel(self, s, e):
        self._s["dialog_result"] = False
        _sigrev_close_with_fade(self._s["win"])

    def ok(self, s, e):
        win = self._s["win"]
        doc = self._s["doc"]
        r_punct = win.FindName("RadRevisionPuntual")
        punctual = r_punct is not None and bool(r_punct.IsChecked)
        self._s["revision_emit_revision_zero"] = punctual
        if punctual and not self._s.get("_sigrev_has_revision_zero", False):
            TaskDialog.Show(
                u"Revisiones",
                u"No hay revisión número 0 en Gestión de revisiones del proyecto.",
            )
            return
        cb_d = win.FindName("CbDescripcion")
        self._s["description"] = (
            unicode(cb_d.SelectedItem)
            if cb_d.SelectedItem is not None
            else ""
        ).strip()
        self._s["dibujo"] = _sigrev_combo_display_to_sheet_value(
            win.FindName("CbDibujo").SelectedItem,
            self._s.get("sigrev_map_dib"),
        )
        self._s["reviso"] = _sigrev_combo_display_to_sheet_value(
            win.FindName("CbReviso").SelectedItem,
            self._s.get("sigrev_map_ing"),
        )
        self._s["aprobo"] = _sigrev_combo_display_to_sheet_value(
            win.FindName("CbAprobo").SelectedItem,
            self._s.get("sigrev_map_ing"),
        )
        fe = unicode(win.FindName("CbFecha").SelectedItem or "")
        self._s["fecha_str"] = fe
        try:
            gd = win.FindName("GridSheets")
            if gd is not None:
                gd.CommitEdit()
        except Exception:
            pass
        self._s["selected_sheets"] = _collect_sheets_checked_in_table(
            doc, self._s.get("sheet_table")
        )
        self._s["dialog_result"] = True
        _sigrev_close_with_fade(win)


def main(__revit__):
    uidoc = __revit__.ActiveUIDocument
    doc = uidoc.Document

    if _try_activate_existing():
        TaskDialog.Show(
            "Revisiones",
            "La herramienta ya está en ejecución.",
        )
        return

    sheets_all = _collect_sheets(doc)
    if not sheets_all:
        forms.alert(u"No hay láminas en el modelo.", title="Revisiones")
        return

    ordered_proj = _revision_ids_ordered_in_project(doc)
    idx_rev0 = _index_revision_project_display_number(doc, ordered_proj, u"0")

    win = XamlReader.Parse(REV_XAML)
    _sigrev_reset_close_guard()

    state = {
        "win": win,
        "doc": doc,
        "sheet_table": None,
        "_sigrev_grid": None,
        "_sigrev_row_delegate": None,
        "_sigrev_syncing_sel_all": False,
        "_sigrev_sel_anchor": None,
        "_sigrev_hdr_chk": None,
        "dialog_result": False,
        "description": "",
        "dibujo": "",
        "reviso": "",
        "aprobo": "",
        "fecha_str": "",
        "selected_sheets": [],
        "_sigrev_has_revision_zero": idx_rev0 >= 0,
        "revision_emit_revision_zero": False,
        "sigrev_map_dib": {},
        "sigrev_map_ing": {},
    }
    handlers = _RevWindowHandlers(state)

    hwnd_owner = None
    try:
        hwnd_owner = revit_main_hwnd(__revit__.Application)
    except Exception:
        pass

    def _sigrev_on_loaded_all(_s, _e):
        try:
            if hwnd_owner:
                from System.Windows.Interop import WindowInteropHelper

                WindowInteropHelper(win).Owner = hwnd_owner
        except Exception:
            pass
        _sigrev_on_win_loaded(win)

    _sigrev_load_logo(win)
    win.Loaded += RoutedEventHandler(_sigrev_on_loaded_all)

    def _sigrev_on_content_rendered(_s, _e):
        # Solo el marco exterior: no forzar SigRevAnimShell aquí (evita cortar el fade-in).
        _sigrev_force_outer_chrome_visible(win)

    win.ContentRendered += EventHandler(_sigrev_on_content_rendered)

    btn_close = win.FindName("BtnClose")
    if btn_close is not None:
        btn_close.Click += RoutedEventHandler(handlers.cancel)
    try:
        from System.Windows.Input import MouseButtonEventHandler

        title_bar = win.FindName("TitleBar")
        if title_bar is not None:
            title_bar.MouseLeftButtonDown += MouseButtonEventHandler(
                lambda _s, e: win.DragMove()
            )
        if btn_close is not None:
            btn_close.MouseLeftButtonDown += MouseButtonEventHandler(
                lambda _s, e: setattr(e, "Handled", True)
            )
    except Exception:
        pass

    try:
        from System.Windows.Input import (
            ApplicationCommands,
            CommandBinding,
            ExecutedRoutedEventHandler,
            KeyBinding,
            Key,
            ModifierKeys,
        )

        win.CommandBindings.Add(
            CommandBinding(
                ApplicationCommands.Close,
                ExecutedRoutedEventHandler(
                    lambda _s, _e: handlers.cancel(_s, _e)
                ),
            )
        )
        win.InputBindings.Add(
            KeyBinding(
                ApplicationCommands.Close,
                Key.Escape,
                getattr(ModifierKeys, "None", 0),
            )
        )
    except Exception:
        pass

    for d in DESCRIPCIONES:
        win.FindName("CbDescripcion").Items.Add(d)
    win.FindName("CbDescripcion").SelectedIndex = 0

    _sigrev_fill_persona_combos(win, state)

    dt_items = win.FindName("CbFecha")
    fecha_opts, fecha_idx_hoy = _dates_dd_mm_yy_options()
    for fx in fecha_opts:
        dt_items.Items.Add(fx)
    dt_items.SelectedIndex = fecha_idx_hoy

    def _sigrev_on_gestionar_personas(_s, _e):
        if GestionarPersonasDialog is None or load_personas_list is None:
            TaskDialog.Show(
                u"Revisiones",
                u"No se pudo cargar el módulo de gestión de personas.",
            )
            return
        oc = ObservableCollection[object]()
        for p in load_personas_list(PERSONAS_FILE):
            oc.Add(p)
        try:
            Directory.CreateDirectory(ISSUES_DIR)
        except Exception:
            pass
        prev_top = None
        try:
            prev_top = win.Topmost
            win.Topmost = False
        except Exception:
            pass
        try:
            GestionarPersonasDialog(
                oc,
                ISSUES_DIR,
                PERSONAS_FILE,
                uidoc=uidoc,
                revit_app=__revit__.Application,
                owner=win,
            )
        except Exception as ex:
            TaskDialog.Show(
                u"Revisiones",
                u"No se pudo abrir el directorio de personas:\n\n{0}".format(str(ex)),
            )
        finally:
            if prev_top is not None:
                try:
                    win.Topmost = prev_top
                except Exception:
                    pass
        _sigrev_fill_persona_combos(win, state)

    btn_gest = win.FindName("BtnGestionarPersonas")
    if btn_gest is not None:
        btn_gest.Click += RoutedEventHandler(_sigrev_on_gestionar_personas)

    if idx_rev0 < 0:
        rp = win.FindName("RadRevisionPuntual")
        if rp is not None:
            try:
                rp.IsEnabled = False
            except Exception:
                pass

    grid_sh = win.FindName("GridSheets")
    state["_sigrev_grid"] = grid_sh
    st_tbl = _build_sheet_selection_table(doc, sheets_all)
    _sigrev_bind_sheet_table(state, st_tbl)
    grid_sh.ItemsSource = st_tbl.DefaultView
    _sigrev_apply_buscar_filter(state)
    _sigrev_sync_buscar_wm(state)
    _sigrev_refresh_estado_sheet(state)
    _sigrev_refresh_nueva_revision_from_ui(state)

    def _sigrev_destino_changed(_snd, _evt):
        _sigrev_refresh_nueva_revision_from_ui(state)

    rad_auto = win.FindName("RadRevAutomatica")
    rad_punt = win.FindName("RadRevisionPuntual")
    if rad_auto is not None:
        rad_auto.Checked += RoutedEventHandler(_sigrev_destino_changed)
    if rad_punt is not None:
        rad_punt.Checked += RoutedEventHandler(_sigrev_destino_changed)

    tb_search = win.FindName("TxtBuscar")
    if tb_search is not None:
        from System.Windows.Controls import TextChangedEventHandler

        tb_search.TextChanged += TextChangedEventHandler(
            lambda s, e: _sigrev_on_buscar_changed(state, s, e)
        )
        tb_search.GotFocus += RoutedEventHandler(
            lambda s, e: _sigrev_sync_buscar_wm(state)
        )
        tb_search.LostFocus += RoutedEventHandler(
            lambda s, e: _sigrev_sync_buscar_wm(state)
        )

    btn_ref = win.FindName("BtnRefrescar")
    if btn_ref is not None:
        btn_ref.Click += RoutedEventHandler(
            lambda s, e: _sigrev_rebind_sheets_grid(state, True)
        )

    grid_sh.Loaded += RoutedEventHandler(
        lambda s, e: _sigrev_on_grid_loaded_sheet(state, s, e)
    )
    try:
        from System.Windows.Input import MouseButtonEventHandler

        grid_sh.PreviewMouseLeftButtonDown += MouseButtonEventHandler(
            lambda s, e: _sigrev_on_grid_preview_mouse(state, s, e)
        )
    except Exception:
        pass

    win.FindName("BtnCancel").Click += RoutedEventHandler(handlers.cancel)
    win.FindName("BtnOk").Click += RoutedEventHandler(handlers.ok)

    def cleanup(s, e):
        _sigrev_detach_row_changed(state)
        _singleton_clear()

    win.Closed += EventHandler(cleanup)

    _singleton_register(win)
    try:
        win.ShowDialog()
    finally:
        _singleton_clear()

    if not state["dialog_result"]:
        return

    sel = state["selected_sheets"]
    if not sel:
        forms.alert(u"Marque al menos una lámina.", title="Revisiones")
        return

    if not state["description"]:
        forms.alert(u"Seleccione una descripción.", title="Revisiones")
        return

    try:
        _parse_dd_mm_yy(state["fecha_str"])
    except Exception:
        forms.alert(
            u"No se interpretó la fecha. Use formato dd.MM.yy.",
            title="Revisiones",
        )
        return

    done, errs = _apply_revision_to_sheets(
        doc,
        sel,
        state["description"],
        state["dibujo"],
        state["reviso"],
        state["aprobo"],
        state["fecha_str"],
        revision_emit_revision_zero=bool(state.get("revision_emit_revision_zero")),
        revit_uiapp=__revit__,
    )

    msg_done = u"Revisiones aplicadas correctamente en {} lámina(s).".format(done)
    if errs.strip():
        msg_done += u"\n\nAdvertencias:\n" + errs
    forms.alert(msg_done, title="Revisiones")
