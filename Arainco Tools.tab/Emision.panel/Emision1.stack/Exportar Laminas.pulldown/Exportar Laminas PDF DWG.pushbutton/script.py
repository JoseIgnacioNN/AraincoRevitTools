# -*- coding: utf-8 -*-
"""
Exportar láminas a PDF y/o DWG con nombre de archivo personalizado por fila.
Carpeta de entrega: ruta completa en el cuadro (tras «Examinar…» se propone
YYYY.MM.DD_ENTREGA bajo la carpeta elegida; el nombre es editable). Dentro: PDF, DWG y opcionalmente listado Excel (láminas seleccionadas, plantilla TemplateListado).
Progreso de exportación: barras ``pyrevit.forms.ProgressBar`` consecutivas — DWG, luego PDF (solo conteo de PDF), luego listado Excel (barra aparte, 1 paso). Acento cian #5BC0DE.
"""

__title__ = "Exportar\nLáminas"
__author__ = "BIMTools"
__doc__ = (
    "Selecciona y exporta láminas (PDF/DWG). Nombre Personalizado: encabezado «Nombre de archivo». "
    "Opcional: listado Excel de las seleccionadas (misma plantilla que «Listado planos Excel»). "
    "Ruta de entrega completa y editable; «Examinar…» la completa con YYYY.MM.DD_ENTREGA (hoy). Subcarpetas PDF y DWG."
)

import imp
import os
import re
import sys

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("System.Windows.Forms")

_pb = os.path.dirname(os.path.abspath(__file__))
_d = _pb
for _ in range(24):
    _sp = os.path.join(_d, "scripts")
    if os.path.isfile(os.path.join(_sp, "revit_wpf_window_position.py")):
        if _sp not in sys.path:
            sys.path.insert(0, _sp)
        break
    _p = os.path.dirname(_d)
    if _p == _d:
        break
    _d = _p
else:
    _sp = os.path.abspath(
        os.path.join(_pb, os.pardir, os.pardir, os.pardir, os.pardir, "scripts")
    )
    if os.path.isdir(_sp) and _sp not in sys.path:
        sys.path.insert(0, _sp)

bimtools_paths = None
try:
    _bimtools_paths_fp = os.path.join(
        os.path.abspath(os.path.join(_pb, os.pardir, os.pardir, os.pardir)),
        "scripts",
        "bimtools_paths.py",
    )
    if os.path.isfile(_bimtools_paths_fp):
        bimtools_paths = imp.load_source(
            "bimtools_paths__ExportarLaminasPDFDWG", _bimtools_paths_fp
        )
    else:
        import bimtools_paths as _bimtools_paths_std

        bimtools_paths = _bimtools_paths_std
    bimtools_paths.set_pushbutton_dir(_pb)
except Exception:
    bimtools_paths = None

# Carga explícita desde la carpeta del botón: evita importar un exportar_laminas_pdf_dwg.py
# viejo en scripts/ o un módulo cacheado en sys.modules.
_EXPORT_MOD_PATH = os.path.join(_pb, "exportar_laminas_pdf_dwg.py")
_EXPORT_MOD_NAME = u"bimtools_exportar_laminas_pdf_dwg__04pushbutton"
for _k in (u"exportar_laminas_pdf_dwg", _EXPORT_MOD_NAME):
    try:
        if _k in sys.modules:
            del sys.modules[_k]
    except Exception:
        pass
if not os.path.isfile(_EXPORT_MOD_PATH):
    raise IOError(
        u"Falta el archivo junto al botón (copia local del módulo): {0}".format(_EXPORT_MOD_PATH)
    )
_export_el = imp.load_source(_EXPORT_MOD_NAME, _EXPORT_MOD_PATH)
build_sheets_datatable = _export_el.build_sheets_datatable
export_sheet_dwg = _export_el.export_sheet_dwg
export_sheet_pdf = _export_el.export_sheet_pdf
sanitize_file_base = _export_el.sanitize_file_base
list_naming_source_options = _export_el.list_naming_source_options
evaluate_naming_recipe = _export_el.evaluate_naming_recipe

_COMPOSER_PATH = os.path.join(_pb, "componer_nombre_lamina_ui.py")
_COMPOSER_MOD_NAME = u"bimtools_componer_nombre_lamina_ui__04pushbutton"
if not os.path.isfile(_COMPOSER_PATH):
    raise IOError(u"Falta el módulo de UI: {0}".format(_COMPOSER_PATH))
_composer_el = imp.load_source(_COMPOSER_MOD_NAME, _COMPOSER_PATH)
show_componer_nombre_dialog = _composer_el.show_componer_nombre_dialog

# Núcleo compartido con 12_ListadoPlanosExcel (misma plantilla y PowerShell).
_LISTADO_PB = os.path.join(os.path.dirname(_pb), u"12_ListadoPlanosExcel.pushbutton")
_TEMPLATE_LISTADO_XLSX = os.path.join(_LISTADO_PB, u"TemplateListado.xlsx")
_LISTADO_CORE_PATH = os.path.join(_LISTADO_PB, u"listado_planos_excel_core.py")
_LISTADO_CORE_NAME = u"bimtools_listado_planos_excel_core__04export_laminas"
_listado_planos_core = None
try:
    if os.path.isfile(_LISTADO_CORE_PATH):
        try:
            if _LISTADO_CORE_NAME in sys.modules:
                del sys.modules[_LISTADO_CORE_NAME]
        except Exception:
            pass
        _listado_planos_core = imp.load_source(_LISTADO_CORE_NAME, _LISTADO_CORE_PATH)
except Exception:
    _listado_planos_core = None

from Autodesk.Revit.DB import ElementId, ViewSheet  # noqa: E402
from Autodesk.Revit.UI import (  # noqa: E402
    TaskDialog,
    TaskDialogCommonButtons,
    TaskDialogResult,
)
from System import EventHandler, Action  # noqa: E402
from System.Windows import RoutedEventHandler, Visibility  # noqa: E402
from System.Windows.Controls import DataGridCellEditEndingEventArgs  # noqa: E402
from System.Windows.Markup import XamlReader  # noqa: E402
import System  # noqa: E402

from revit_wpf_window_position import (  # noqa: E402
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""

doc = __revit__.ActiveUIDocument.Document  # noqa: F821
uidoc = __revit__.ActiveUIDocument  # noqa: F821

try:
    from join_geometry_concrete_vista import (
        _BloquearComandosRevit,
        _pbar_exit_safe,
        _pbar_step,
    )
except Exception:
    _BloquearComandosRevit = None
    _pbar_exit_safe = None
    _pbar_step = None

try:
    from pyrevit import forms as _pyrevit_forms_export_lam
except Exception:
    _pyrevit_forms_export_lam = None

# Acento del formulario / barra pyRevit (cian #5BC0DE).
_EXP_LAM_PYREVIT_ACCENT_RGB = (91, 192, 222)

# Texto de la barra superior (pyRevit ProgressBar): «Arainco - Exportando … X/Y».
_EXP_LAM_PBAR_DWG = u"Arainco - Exportando DWG"
_EXP_LAM_PBAR_PDF = u"Arainco - Exportando PDF"
_EXP_LAM_PBAR_LISTADO = u"Arainco - Exportando listado Excel"


def _export_laminas_pbar_initial_title(base, total):
    try:
        t = int(total)
    except Exception:
        t = 0
    if t < 1:
        t = 1
    return u"{} 0/{}".format(base, t)


def _export_laminas_pbar_step(pb, current_index, count, base_title):
    """Avance de ProgressBar; título ``base X/Y`` con un solo espacio (formato Arainco)."""
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


def _export_laminas_pbar_start(title, count):
    """
    ``forms.ProgressBar`` de pyRevit (misma barra superior que Armado Muros Nodo).
    Asigna ``pyRevitAccentBrush`` al cian de la herramienta cuando el runtime lo permite.
    """
    if _pyrevit_forms_export_lam is None or count is None or int(count) < 1:
        return None
    try:
        pb = _pyrevit_forms_export_lam.ProgressBar(
            title=title,
            cancellable=False,
        )
        try:
            from System.Windows.Media import Color, SolidColorBrush

            r, g, b = _EXP_LAM_PYREVIT_ACCENT_RGB
            pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(Color.FromRgb(r, g, b))
        except Exception:
            pass
        return pb
    except Exception:
        return None


if _pbar_step is None:

    def _pbar_step(pb, current_index, count, base_title):
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
            pb.title = u"{}  {}/{}".format(base_title, i, c)
        except Exception:
            pass


if _pbar_exit_safe is None:

    def _pbar_exit_safe(pb, ok):
        if ok and pb is not None:
            try:
                pb.__exit__(None, None, None)
            except Exception:
                pass


_APPDOMAIN_WINDOW_KEY = u"BIMTools.ExportarLaminasPDFDWG.ActiveWindow"

# TaskDialog: con TitleAutoPrefix (por defecto), Revit antepone el nombre del botón al título.
_TASK_DLG_EXPORT_LAM_TITLE = u"Exportar Láminas"

# Setup DWG del proyecto (sin UI): mismo criterio que Revit «Default».
_DWG_EXPORT_SETUP_NAME = u"Default"

# Apertura/cierre: mismo ritmo que Fundación aislada (ScaleTransform + opacidad).
_EXPORT_LAM_CHROME_MS = 180
_WPF_STORYBOARD_DUR_STR = u"0:0:{0:.2f}".format(_EXPORT_LAM_CHROME_MS / 1000.0)


def _nombre_carpeta_entrega_por_defecto():
    """Nombre sugerido para la carpeta de entrega: fecha local YYYY.MM.DD_ENTREGA."""
    from datetime import date

    return date.today().strftime("%Y.%m.%d") + u"_ENTREGA"


def _es_nombre_carpeta_entrega_estandar(basename):
    """True si el último segmento cumple YYYY.MM.DD_ENTREGA (evita duplicar al re-examinar)."""
    try:
        s = unicode(basename)
    except Exception:
        return False
    if not s.endswith(u"_ENTREGA"):
        return False
    date_part = s[:-8]
    parts = date_part.split(u".")
    if len(parts) != 3:
        return False
    if len(parts[0]) != 4 or len(parts[1]) != 2 or len(parts[2]) != 2:
        return False
    for p in parts:
        if not p.isdigit():
            return False
    return True


def _load_bimtools_logo_into_window(win):
    """Misma resolución de logo que en el resto de herramientas BIMTools (`bimtools_paths`)."""
    try:
        if bimtools_paths is None:
            return

        img_ctrl = win.FindName(u"ImgLogo")
        if not img_ctrl:
            return
        bmp = bimtools_paths.load_logo_bitmap_image()
        if bmp is None:
            return
        img_ctrl.Source = bmp
        try:
            win.Icon = bmp
        except Exception:
            pass
    except Exception:
        pass


XAML = (
    u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="ExpLamWin"
    Title="Exportar Láminas"
    Height="810" Width="1040" MinHeight="640" MinWidth="800"
    Background="Transparent"
    AllowsTransparency="True"
    WindowStyle="None"
    ResizeMode="NoResize"
    WindowStartupLocation="Manual"
    Topmost="True"
    UseLayoutRounding="True"
    FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
    <Storyboard x:Key="ExpLamOpenGrowStoryboard">
      <DoubleAnimation Storyboard.TargetName="ExpLamRootScale" Storyboard.TargetProperty="ScaleX"
                       From="0" To="1" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="ExpLamRootScale" Storyboard.TargetProperty="ScaleY"
                       From="0" To="1" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
      <DoubleAnimation Storyboard.TargetName="ExpLamWin" Storyboard.TargetProperty="Opacity"
                       From="0" To="1" Duration="__WPF_STORYBOARD_DUR__" FillBehavior="HoldEnd">
        <DoubleAnimation.EasingFunction>
          <QuadraticEase EasingMode="EaseOut"/>
        </DoubleAnimation.EasingFunction>
      </DoubleAnimation>
    </Storyboard>
"""
    + BIMTOOLS_DARK_STYLES_XML
    + u"""
    <Style x:Key="StepBadge" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#5BC0DE"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
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
    <Style x:Key="BtnTertiary" TargetType="Button" BasedOn="{StaticResource BtnSelectOutline}">
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="Background" Value="#0D1829"/>
      <Setter Property="Foreground" Value="#B8D4E5"/>
      <Setter Property="BorderBrush" Value="#355973"/>
    </Style>
    <Style x:Key="ExpFmtToggle" TargetType="CheckBox">
      <Setter Property="Foreground" Value="#B8D4E5"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Border x:Name="Bd" Background="#0D1829" BorderBrush="#355973" BorderThickness="1" CornerRadius="5" Padding="8,5" MinHeight="30" HorizontalAlignment="Stretch" SnapsToDevicePixels="True">
              <Grid HorizontalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Box" Grid.Column="0" Width="14" Height="14" Background="#050E18" BorderBrush="#355973" BorderThickness="1" CornerRadius="2" Margin="0,0,6,0" VerticalAlignment="Center">
                  <Path x:Name="Glyph" Data="M 2.5,7 L 5,9.5 L 10,3.5" Stroke="#5BC0DE" StrokeThickness="1.5" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Visibility="Collapsed" SnapsToDevicePixels="True"/>
                </Border>
                <ContentPresenter Grid.Column="1" VerticalAlignment="Center" RecognizesAccessKey="True"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Glyph" Property="Visibility" Value="Visible"/>
                <Setter TargetName="Box" Property="BorderBrush" Value="#5BC0DE"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#0F2035"/>
                <Setter TargetName="Bd" Property="BorderBrush" Value="#5BC0DE"/>
                <Setter Property="Foreground" Value="#E8F4F8"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.45"/>
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
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#050E18"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#284760"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="FontSize" Value="11"/>
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
      <Setter Property="TextAlignment" Value="Left"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="DgTbCenter" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="TextAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="DgEditLeft" TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
      <Setter Property="TextAlignment" Value="Left"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="DgTbLeftPadded" TargetType="TextBlock" BasedOn="{StaticResource DgTbLeft}">
      <Setter Property="Margin" Value="12,0,8,0"/>
    </Style>
    <Style x:Key="DgEditLeftPadded" TargetType="TextBox" BasedOn="{StaticResource DgEditLeft}">
      <Setter Property="Padding" Value="18,6,8,6"/>
    </Style>
    <Style x:Key="DgHdrCenter" TargetType="DataGridColumnHeader" BasedOn="{StaticResource {x:Type DataGridColumnHeader}}">
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="DgHdrLeft" TargetType="DataGridColumnHeader" BasedOn="{StaticResource {x:Type DataGridColumnHeader}}">
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
    </Style>
    <Style x:Key="BtnNombreArchivoHeader" TargetType="Button">
      <Setter Property="Background" Value="#0D1829"/>
      <Setter Property="Foreground" Value="#C8E4EF"/>
      <Setter Property="BorderBrush" Value="#355973"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Padding" Value="12,10"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{TemplateBinding Background}" CornerRadius="5"
                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#0F2035"/>
                <Setter TargetName="Bd" Property="BorderBrush" Value="#5BC0DE"/>
                <Setter Property="Foreground" Value="#E8F4F8"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#1B3D5C"/>
                <Setter TargetName="Bd" Property="BorderBrush" Value="#5BC0DE"/>
                <Setter Property="Foreground" Value="#FFFFFF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="DgHdrNombreArchivo" TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrLeft}">
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Padding" Value="4,6"/>
      <Setter Property="Background" Value="Transparent"/>
    </Style>
    <!-- Misma barra que BIMTools (flechas rellenas, ~18px): la versión ExpLam 10px + trazos se leía mal. -->
    <Style x:Key="ExpLamScrollBarDark" TargetType="ScrollBar" BasedOn="{StaticResource BimToolsScrollBarDark}"/>
  </Window.Resources>
  <Border x:Name="ExpLamRootChrome" CornerRadius="8" Background="#0E1B32" Padding="14"
          BorderBrush="#5BC0DE" BorderThickness="1" ClipToBounds="True" RenderTransformOrigin="0,0">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="16" ShadowDepth="0" Opacity="0.35"/>
    </Border.Effect>
    <Border.RenderTransform>
      <ScaleTransform x:Name="ExpLamRootScale" ScaleX="0" ScaleY="0"/>
    </Border.RenderTransform>
    <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
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
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <Image x:Name="ImgLogo" Width="40" Height="40"
                 Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center" RenderOptions.BitmapScalingMode="HighQuality"/>
          <TextBlock Text="Exportar Láminas" FontSize="14" FontWeight="Bold"
                     Foreground="#FFFFFF" TextWrapping="NoWrap" VerticalAlignment="Center"
                     TextAlignment="Left"/>
        </StackPanel>
        <Button x:Name="BtnClose" Grid.Column="2"
                Style="{StaticResource BtnCloseX_MinimalNoBg}"
                VerticalAlignment="Center" HorizontalAlignment="Right" ToolTip="Cerrar"/>
      </Grid>
    </Border>
    <Border Grid.Row="1" Style="{StaticResource PanelInset}" Margin="0,0,0,10" Padding="10,8">
      <Grid VerticalAlignment="Center">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid Grid.Column="0" Margin="0,0,10,0">
          <Border Background="#050E18" BorderBrush="#355973" BorderThickness="1" CornerRadius="5" Padding="0" MinHeight="32">
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
                         ToolTip="Filtrar por número o nombre de lámina (lista principal)"/>
                <TextBlock x:Name="TxtBuscarWatermark" Text="Buscar" IsHitTestVisible="False"
                           Foreground="#5C7A8F" FontSize="12" VerticalAlignment="Center" Margin="0,0,10,0"/>
              </Grid>
            </Grid>
          </Border>
        </Grid>
        <CheckBox x:Name="ChkPdf" Grid.Column="1" Style="{StaticResource ExpFmtToggle}" Content="PDF" IsChecked="True" Margin="0,0,8,0"
                  ToolTip="Exportar a PDF (subcarpeta /PDF dentro de la ruta de entrega del cuadro inferior)."/>
        <CheckBox x:Name="ChkDwg" Grid.Column="2" Style="{StaticResource ExpFmtToggle}" Content="DWG" IsChecked="True" Margin="0,0,8,0"
                  ToolTip="Exportar a DWG con setup «Default» del proyecto (subcarpeta /DWG). MergedViews: un solo archivo."/>
        <CheckBox x:Name="ChkListadoPlanos" Grid.Column="3" Style="{StaticResource ExpFmtToggle}" Content="Listado" IsChecked="True" Margin="0,0,8,0"
                  ToolTip="Excel en la carpeta de entrega: solo las láminas seleccionadas (plantilla TemplateListado, requiere Excel)."/>
        <Button x:Name="BtnRefrescar" Grid.Column="4" Style="{StaticResource BtnGhost}" Margin="0,0,0,0" Padding="8,5" FontSize="10"
                ToolTip="Actualizar la lista de láminas desde el proyecto">
          <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
            <TextBlock Text="&#xE72C;" FontFamily="Segoe MDL2 Assets" FontSize="13"
                       VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBlock Text="Actualizar" VerticalAlignment="Center" FontSize="10" FontWeight="SemiBold"/>
          </StackPanel>
        </Button>
      </Grid>
    </Border>
    <Border Grid.Row="2" Background="#040A12" BorderBrush="#1E3F55" BorderThickness="1" CornerRadius="8" Padding="0" Margin="0,0,0,12">
      <DataGrid x:Name="GridSheets" MinHeight="396" AutoGenerateColumns="False" CanUserAddRows="False"
                RowHeaderWidth="0" SelectionMode="Extended" SelectionUnit="FullRow"
                AlternationCount="2"
                ClipboardCopyMode="ExcludeHeader" HeadersVisibility="Column">
        <DataGrid.CellStyle>
          <Style TargetType="DataGridCell" BasedOn="{StaticResource GridCellPadding}"/>
        </DataGrid.CellStyle>
        <DataGrid.Columns>
          <DataGridCheckBoxColumn Header="" Binding="{Binding Sel, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="44" CanUserSort="False">
            <DataGridCheckBoxColumn.HeaderTemplate>
              <DataTemplate>
                <CheckBox Tag="HdrSelectAll" Style="{StaticResource ChkSheetSel}" HorizontalAlignment="Center"
                          VerticalAlignment="Center" IsThreeState="True"
                          ToolTip="Marcar o anular todas las láminas visibles en la lista (respeta el filtro Buscar)."/>
              </DataTemplate>
            </DataGridCheckBoxColumn.HeaderTemplate>
            <DataGridCheckBoxColumn.HeaderStyle>
              <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrCenter}">
                <Setter Property="ToolTip" Value="Filas: un clic en la casilla marca o desmarca; Mayús+clic entre dos filas aplica el mismo estado al rango visible. Cabecera: marcar/anular todas las visibles."/>
              </Style>
            </DataGridCheckBoxColumn.HeaderStyle>
            <DataGridCheckBoxColumn.ElementStyle>
              <Style TargetType="CheckBox" BasedOn="{StaticResource ChkSheetSel}"/>
            </DataGridCheckBoxColumn.ElementStyle>
            <DataGridCheckBoxColumn.EditingElementStyle>
              <Style TargetType="CheckBox" BasedOn="{StaticResource ChkSheetSel}"/>
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
          <DataGridTextColumn Header="Nombre" Binding="{Binding SheetName}" Width="*" MinWidth="200" IsReadOnly="True">
            <DataGridTextColumn.HeaderStyle>
              <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrLeft}"/>
            </DataGridTextColumn.HeaderStyle>
            <DataGridTextColumn.ElementStyle>
              <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbLeftPadded}"/>
            </DataGridTextColumn.ElementStyle>
          </DataGridTextColumn>
          <DataGridTextColumn Header="Revisión" Binding="{Binding Revision}" Width="88" IsReadOnly="True">
            <DataGridTextColumn.HeaderStyle>
              <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrCenter}"/>
            </DataGridTextColumn.HeaderStyle>
            <DataGridTextColumn.ElementStyle>
              <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbCenter}"/>
            </DataGridTextColumn.ElementStyle>
          </DataGridTextColumn>
          <DataGridTextColumn Binding="{Binding CustomName, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="*" MinWidth="220" CanUserSort="False">
            <DataGridTextColumn.HeaderTemplate>
              <DataTemplate>
                <Button Tag="HdrComponer" Style="{StaticResource BtnNombreArchivoHeader}"
                        HorizontalAlignment="Stretch" MinHeight="36"
                        Content="Nombre de archivo"
                        ToolTip="Pulse el encabezado para abrir Nombre Personalizado (componer nombre por parámetros de lámina; todas las filas)."/>
              </DataTemplate>
            </DataGridTextColumn.HeaderTemplate>
            <DataGridTextColumn.HeaderStyle>
              <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrNombreArchivo}"/>
            </DataGridTextColumn.HeaderStyle>
            <DataGridTextColumn.ElementStyle>
              <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbLeftPadded}"/>
            </DataGridTextColumn.ElementStyle>
            <DataGridTextColumn.EditingElementStyle>
              <Style TargetType="TextBox" BasedOn="{StaticResource DgEditLeftPadded}"/>
            </DataGridTextColumn.EditingElementStyle>
          </DataGridTextColumn>
          <DataGridTextColumn Header="" Binding="{Binding IdInt}" Width="0" MinWidth="0" MaxWidth="0" IsReadOnly="True">
            <DataGridTextColumn.ElementStyle>
              <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbLeft}"/>
            </DataGridTextColumn.ElementStyle>
          </DataGridTextColumn>
        </DataGrid.Columns>
      </DataGrid>
    </Border>
    <TextBlock Grid.Row="3" Text="Carpeta de salida" Style="{StaticResource StepBadge}"/>
    <Border Grid.Row="4" Background="Transparent" Margin="0,0,0,0">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Border Grid.Row="0" Style="{StaticResource PanelInset}" Margin="0,0,0,10">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="TxtCarpeta" Grid.Column="0" Margin="0,0,12,0"/>
            <Button x:Name="BtnCarpeta" Grid.Column="1" Content="Examinar…" Style="{StaticResource BtnSelectOutline}"
                    MinWidth="132" MinHeight="38" Padding="18,9" FontSize="12"/>
          </Grid>
        </Border>
        <Border Grid.Row="1" Background="#050E18" BorderBrush="#1E3F55" BorderThickness="1" CornerRadius="6"
                Padding="14,12" Margin="0,0,0,0">
          <Border.Effect>
            <DropShadowEffect BlurRadius="18" ShadowDepth="0" Opacity="0.2" Color="#000000"/>
          </Border.Effect>
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" VerticalAlignment="Center" Orientation="Vertical">
              <TextBlock x:Name="TxtEstado" Foreground="#E8F4F8" FontSize="12" FontWeight="SemiBold" TextWrapping="Wrap"/>
              <TextBlock Foreground="#6B94AA" FontSize="10" Margin="0,4,0,0" Text="Estado del listado · listo para exportar"/>
            </StackPanel>
            <Button x:Name="BtnExportar" Grid.Column="1" Content="Exportar" Style="{StaticResource BtnPrimary}" MinWidth="132" MinHeight="38"/>
          </Grid>
        </Border>
      </Grid>
    </Border>
    </Grid>
  </Border>
</Window>
"""
).replace(u"__WPF_STORYBOARD_DUR__", _WPF_STORYBOARD_DUR_STR)


def _clear_appdomain_window():
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _apply_bimtools_scrollbars_below(root_visual, resources_owner):
    """Fuerza estilo oscuro en cada ScrollBar bajo root_visual (ExpLamScrollBarDark → BimToolsScrollBarDark)."""
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


def _schedule_export_laminas_scrollbar_styles(win, grid):
    """Varias pasadas: el DataGrid recrea el ScrollViewer al medir filas."""
    try:
        from System.Windows.Threading import DispatcherPriority

        def _go():
            _apply_bimtools_scrollbars_below(grid, win)
            _apply_bimtools_scrollbars_below(win, win)

        _go()
        win.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, Action(_go))
    except Exception:
        try:
            _apply_bimtools_scrollbars_below(grid, win)
            _apply_bimtools_scrollbars_below(win, win)
        except Exception:
            pass


def _get_active_tool_window():
    try:
        win = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        _ = win.Title
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            _clear_appdomain_window()
            return None
    except Exception:
        _clear_appdomain_window()
        return None
    return win


def _run_revit_taskdialogs_above_wpf(wpf_win, callback):
    """
    El formulario usa Topmost=True; los TaskDialog de Revit quedan detrás si no se
    desactiva temporalmente (mismo criterio que ``_task_dialog_safe``).
    """
    top = None
    if wpf_win is not None:
        try:
            top = wpf_win.Topmost
            wpf_win.Topmost = False
        except Exception:
            top = None
    try:
        callback()
    finally:
        if wpf_win is not None and top is not None:
            try:
                wpf_win.Topmost = top
            except Exception:
                pass


def _export_laminas_td_ok(wpf_win, main_instruction):
    """TaskDialog OK con título Arainco; evita el prefijo automático del botón pyRevit."""
    def _cb():
        td = TaskDialog(_TASK_DLG_EXPORT_LAM_TITLE)
        try:
            td.TitleAutoPrefix = False
        except Exception:
            pass
        td.MainInstruction = main_instruction
        td.CommonButtons = TaskDialogCommonButtons.Ok
        td.DefaultButton = TaskDialogResult.Ok
        td.Show()

    _run_revit_taskdialogs_above_wpf(wpf_win, _cb)


def _task_dialog_safe(title, msg, wpf_win=None):
    def _cb():
        td = TaskDialog(title)
        try:
            td.TitleAutoPrefix = False
        except Exception:
            pass
        td.MainInstruction = msg
        td.CommonButtons = TaskDialogCommonButtons.Ok
        td.DefaultButton = TaskDialogResult.Ok
        td.Show()

    _run_revit_taskdialogs_above_wpf(wpf_win, _cb)


class ExportarLaminasWindow(object):
    def __init__(self):
        self._open_grow_storyboard_started = False
        self._is_closing_with_fade = False
        self._syncing_select_all = False
        self._chk_select_all = None
        self._sel_anchor_idx = None
        self._revit_cmd_blocker = None
        self._win = XamlReader.Parse(XAML)
        _load_bimtools_logo_into_window(self._win)
        self._grid = self._win.FindName("GridSheets")
        self._txt_buscar = self._win.FindName("TxtBuscar")
        self._txt_buscar_watermark = self._win.FindName("TxtBuscarWatermark")
        self._txt_carpeta = self._win.FindName("TxtCarpeta")
        self._txt_estado = self._win.FindName("TxtEstado")
        self._chk_pdf = self._win.FindName("ChkPdf")
        self._chk_dwg = self._win.FindName("ChkDwg")
        self._chk_listado_plan = self._win.FindName("ChkListadoPlanos")
        self._btn_export = self._win.FindName("BtnExportar")
        # Capturamos en self: IronPython 2 + pyRevit no expone globales del módulo
        # en métodos de clase llamados desde eventos WPF asíncronos. self siempre accesible.
        self._doc = doc
        self._export_sheet_pdf = export_sheet_pdf
        self._export_sheet_dwg = export_sheet_dwg
        self._sanitize_file_base = sanitize_file_base
        self._table = build_sheets_datatable(doc)
        self._table.RowChanged += self._on_table_row_changed
        self._grid.ItemsSource = self._table.DefaultView
        self._grid.Loaded += RoutedEventHandler(self._on_grid_loaded)
        self._win.Loaded += RoutedEventHandler(self._on_win_loaded_scrollbars)
        self._grid.CellEditEnding += EventHandler[DataGridCellEditEndingEventArgs](
            self._on_cell_edit_ending
        )
        try:
            from System.Windows.Input import MouseButtonEventHandler

            self._grid.PreviewMouseLeftButtonDown += MouseButtonEventHandler(
                self._on_grid_sel_preview_mouse_left_button_down
            )
        except Exception:
            pass
        self._txt_buscar.TextChanged += self._on_buscar_changed
        self._txt_buscar.GotFocus += RoutedEventHandler(
            lambda s, e: self._sync_buscar_watermark()
        )
        self._txt_buscar.LostFocus += RoutedEventHandler(
            lambda s, e: self._sync_buscar_watermark()
        )
        self._sync_buscar_watermark()
        self._win.FindName("BtnRefrescar").Click += RoutedEventHandler(self._on_refrescar)
        self._win.FindName("BtnCarpeta").Click += RoutedEventHandler(self._on_carpeta)
        self._btn_export.Click += RoutedEventHandler(self._on_exportar)

        btn_close = self._win.FindName("BtnClose")
        if btn_close is not None:
            btn_close.Click += RoutedEventHandler(lambda s, e: self._close_with_fade())
        try:
            from System.Windows.Input import MouseButtonEventHandler

            title_bar = self._win.FindName("TitleBar")
            if title_bar is not None:
                title_bar.MouseLeftButtonDown += MouseButtonEventHandler(
                    lambda s, e: self._win.DragMove()
                )
            if btn_close is not None:
                btn_close.MouseLeftButtonDown += MouseButtonEventHandler(
                    lambda s, e: setattr(e, "Handled", True)
                )
        except Exception:
            pass

        self._wire_close_commands()

        def _closed(s, e):
            try:
                self._release_revit_cmd_blocker()
            except Exception:
                pass
            _clear_appdomain_window()

        self._win.Closed += EventHandler(_closed)
        self._refresh_estado()

    def _wire_close_commands(self):
        try:
            from System.Windows.Input import (
                ApplicationCommands,
                CommandBinding,
                ExecutedRoutedEventHandler,
                KeyBinding,
                Key,
                ModifierKeys,
            )

            self._win.CommandBindings.Add(
                CommandBinding(
                    ApplicationCommands.Close,
                    ExecutedRoutedEventHandler(lambda s, e: self._close_with_fade()),
                )
            )
            self._win.InputBindings.Add(
                KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None)
            )
        except Exception:
            pass

    def _ensure_revit_cmd_blocker(self):
        """Inhabilita la ventana principal de Revit mientras el formulario está abierto."""
        if self._revit_cmd_blocker is not None:
            return
        if _BloquearComandosRevit is None:
            return
        try:
            b = _BloquearComandosRevit(__revit__)
            b.__enter__()
            self._revit_cmd_blocker = b
        except Exception:
            self._revit_cmd_blocker = None

    def _release_revit_cmd_blocker(self):
        b = self._revit_cmd_blocker
        self._revit_cmd_blocker = None
        if b is not None:
            try:
                b.__exit__(None, None, None)
            except Exception:
                pass

    def _begin_open_grow_storyboard(self):
        if self._open_grow_storyboard_started:
            return
        self._open_grow_storyboard_started = True
        try:
            from System import TimeSpan
            from System.Windows import Duration

            sc = self._win.FindName("ExpLamRootScale")
            if sc is not None:
                sc.ScaleX = 0.0
                sc.ScaleY = 0.0
            sb = self._win.TryFindResource("ExpLamOpenGrowStoryboard")
            if sb is None:
                if sc is not None:
                    sc.ScaleX = sc.ScaleY = 1.0
                self._win.Opacity = 1.0
                return
            dur = Duration(TimeSpan.FromMilliseconds(float(_EXPORT_LAM_CHROME_MS)))
            try:
                for i in range(int(sb.Children.Count)):
                    sb.Children[i].Duration = dur
            except Exception:
                pass
            sb.Begin(self._win, True)
        except Exception:
            try:
                self._win.Opacity = 1.0
                sc = self._win.FindName("ExpLamRootScale")
                if sc is not None:
                    sc.ScaleX = sc.ScaleY = 1.0
            except Exception:
                pass

    def _close_with_fade(self):
        if self._is_closing_with_fade:
            return
        self._is_closing_with_fade = True
        try:
            from System import TimeSpan, EventHandler
            from System.Windows import Duration
            from System.Windows.Media import ScaleTransform
            from System.Windows.Media.Animation import DoubleAnimation, QuadraticEase, EasingMode

            sc = self._win.FindName("ExpLamRootScale")
            dur = Duration(TimeSpan.FromMilliseconds(float(_EXPORT_LAM_CHROME_MS)))
            ease_in = QuadraticEase()
            ease_in.EasingMode = EasingMode.EaseIn

            def _da(f0, f1):
                a = DoubleAnimation()
                a.From = float(f0)
                a.To = float(f1)
                a.Duration = dur
                a.EasingFunction = ease_in
                return a

            try:
                sx0 = float(sc.ScaleX) if sc is not None else 1.0
                sy0 = float(sc.ScaleY) if sc is not None else 1.0
            except Exception:
                sx0 = sy0 = 1.0
            try:
                op0 = float(self._win.Opacity)
            except Exception:
                op0 = 1.0

            op_anim = _da(op0, 0.0)
            ax = _da(sx0, 0.0)
            ay = _da(sy0, 0.0)

            def _done(sender, args):
                try:
                    self._win.Close()
                except Exception:
                    pass

            op_anim.Completed += EventHandler(_done)
            if sc is not None:
                sc.BeginAnimation(ScaleTransform.ScaleXProperty, ax)
                sc.BeginAnimation(ScaleTransform.ScaleYProperty, ay)
            self._win.BeginAnimation(self._win.OpacityProperty, op_anim)
        except Exception:
            try:
                self._win.Close()
            except Exception:
                pass
            self._is_closing_with_fade = False

    def _on_grid_loaded(self, sender, args):
        try:
            if self._grid.Columns.Count > 0:
                c = self._grid.Columns[self._grid.Columns.Count - 1]
                c.Visibility = Visibility.Collapsed
        except Exception:
            pass
        try:
            _schedule_export_laminas_scrollbar_styles(self._win, self._grid)
        except Exception:
            pass
        try:
            from System.Windows.Threading import DispatcherPriority

            def _wire_hdr():
                try:
                    self._wire_custom_name_header()
                    self._wire_select_all_header()
                    self._sync_select_all_header()
                except Exception:
                    pass

            _wire_hdr()
            self._grid.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(_wire_hdr))
            self._grid.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_wire_hdr))
        except Exception:
            try:
                self._wire_custom_name_header()
                self._wire_select_all_header()
                self._sync_select_all_header()
            except Exception:
                pass

    def _wire_custom_name_header(self):
        from System.Windows.Media import VisualTreeHelper
        from System.Windows.Controls import Button

        def walk_find(compositor):
            try:
                n = VisualTreeHelper.GetChildrenCount(compositor)
            except Exception:
                return None
            for i in range(n):
                try:
                    ch = VisualTreeHelper.GetChild(compositor, i)
                except Exception:
                    continue
                try:
                    if isinstance(ch, Button):
                        tg = ch.Tag
                        if tg is not None and unicode(str(tg)) == u"HdrComponer":
                            return ch
                except Exception:
                    pass
                found = walk_find(ch)
                if found is not None:
                    return found
            return None

        btn = walk_find(self._grid)
        if btn is None:
            return
        if not hasattr(self, u"_hdr_componer_handler"):
            self._hdr_componer_handler = RoutedEventHandler(self._on_componer_nombre)
        try:
            btn.Click -= self._hdr_componer_handler
        except Exception:
            pass
        btn.Click += self._hdr_componer_handler

    def _wire_select_all_header(self):
        from System.Windows.Media import VisualTreeHelper
        from System.Windows.Controls import CheckBox

        def walk_find(compositor):
            try:
                n = VisualTreeHelper.GetChildrenCount(compositor)
            except Exception:
                return None
            for i in range(n):
                try:
                    ch = VisualTreeHelper.GetChild(compositor, i)
                except Exception:
                    continue
                try:
                    if isinstance(ch, CheckBox):
                        tg = ch.Tag
                        if tg is not None and unicode(str(tg)) == u"HdrSelectAll":
                            return ch
                except Exception:
                    pass
                found = walk_find(ch)
                if found is not None:
                    return found
            return None

        chk = walk_find(self._grid)
        if chk is None:
            return
        self._chk_select_all = chk
        if not hasattr(self, u"_select_all_header_click_handler"):
            self._select_all_header_click_handler = RoutedEventHandler(
                self._on_select_all_header_click
            )
        try:
            chk.Click -= self._select_all_header_click_handler
        except Exception:
            pass
        chk.Click += self._select_all_header_click_handler

    def _on_select_all_header_click(self, sender, args):
        if getattr(self, "_syncing_select_all", False):
            return
        try:
            args.Handled = True
        except Exception:
            pass
        dv = self._table.DefaultView
        n = dv.Count
        if n == 0:
            return
        n_sel = 0
        for i in range(n):
            try:
                if dv[i].Row[u"Sel"]:
                    n_sel += 1
            except Exception:
                pass
        new_val = True
        if n_sel == n:
            new_val = False
        self._syncing_select_all = True
        try:
            for i in range(n):
                try:
                    dv[i].Row[u"Sel"] = new_val
                except Exception:
                    pass
        finally:
            self._syncing_select_all = False
        self._refresh_estado()

    def _sync_select_all_header(self):
        chk = getattr(self, "_chk_select_all", None)
        if chk is None:
            try:
                self._wire_select_all_header()
            except Exception:
                pass
            chk = getattr(self, "_chk_select_all", None)
        if chk is None:
            return
        dv = self._table.DefaultView
        n = dv.Count
        self._syncing_select_all = True
        try:
            try:
                chk.IsThreeState = True
            except Exception:
                pass
            if n == 0:
                chk.IsChecked = False
                return
            n_sel = 0
            for i in range(n):
                try:
                    if dv[i].Row[u"Sel"]:
                        n_sel += 1
                except Exception:
                    pass
            if n_sel == 0:
                chk.IsChecked = False
            elif n_sel == n:
                chk.IsChecked = True
            else:
                chk.IsChecked = None
        finally:
            self._syncing_select_all = False

    def _on_win_loaded_scrollbars(self, sender, args):
        try:
            _schedule_export_laminas_scrollbar_styles(self._win, self._grid)
        except Exception:
            pass
        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            self._win.Dispatcher.BeginInvoke(
                Action(self._begin_open_grow_storyboard),
                DispatcherPriority.Loaded,
            )
        except Exception:
            try:
                self._begin_open_grow_storyboard()
            except Exception:
                pass

    def _on_cell_edit_ending(self, sender, args):
        self._refresh_estado()

    def _on_table_row_changed(self, sender, args):
        self._refresh_estado()

    def _refresh_estado(self):
        n = self._table.Rows.Count
        ns = 0
        for i in range(n):
            try:
                if self._table.Rows[i][u"Sel"]:
                    ns += 1
            except Exception:
                pass
        self._txt_estado.Text = u"{0} láminas  |  {1} seleccionadas".format(n, ns)
        if not getattr(self, "_syncing_select_all", False):
            self._sync_select_all_header()

    def _sync_buscar_watermark(self):
        wm = getattr(self, "_txt_buscar_watermark", None)
        tb = getattr(self, "_txt_buscar", None)
        if wm is None or tb is None:
            return
        try:
            t = tb.Text
            t = unicode(t).strip() if t is not None else u""
        except Exception:
            t = u""
        try:
            focused = bool(tb.IsFocused)
        except Exception:
            focused = False
        show = (not t) and (not focused)
        try:
            wm.Visibility = Visibility.Visible if show else Visibility.Collapsed
        except Exception:
            pass

    def _on_buscar_changed(self, sender, args):
        self._sync_buscar_watermark()
        dv = self._table.DefaultView
        try:
            t = self._txt_buscar.Text
            t = unicode(t).strip() if t is not None else u""
        except Exception:
            t = u""
        if not t:
            dv.RowFilter = u""
        else:
            esc = t.replace(u"'", u"''")
            dv.RowFilter = u"[SheetNumber] LIKE '%{0}%' OR [SheetName] LIKE '%{0}%'".format(esc)
        self._sync_select_all_header()

    def _on_refrescar(self, sender, args):
        self._chk_select_all = None
        self._sel_anchor_idx = None
        try:
            self._table.RowChanged -= self._on_table_row_changed
        except Exception:
            pass
        self._table = build_sheets_datatable(doc)
        self._table.RowChanged += self._on_table_row_changed
        self._grid.ItemsSource = self._table.DefaultView
        self._txt_buscar.Text = u""
        self._refresh_estado()
        try:
            from System.Windows.Threading import DispatcherPriority

            def _wire_hdr():
                try:
                    self._wire_custom_name_header()
                    self._wire_select_all_header()
                    self._sync_select_all_header()
                except Exception:
                    pass

            self._grid.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(_wire_hdr))
            self._grid.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_wire_hdr))
        except Exception:
            try:
                self._wire_custom_name_header()
                self._wire_select_all_header()
                self._sync_select_all_header()
            except Exception:
                pass

    def _on_componer_nombre(self, sender, args):
        try:
            args.Handled = True
        except Exception:
            pass
        try:
            show_componer_nombre_dialog(
                self._win,
                self._doc,
                self._table,
                list_naming_source_options,
                evaluate_naming_recipe,
            )
        except Exception as _ex:
            try:
                TaskDialog.Show(
                    u"Nombre Personalizado",
                    u"Error al abrir Nombre Personalizado:\n\n{0}".format(unicode(str(_ex))),
                )
            except Exception:
                pass
        self._refresh_estado()

    def _on_carpeta(self, sender, args):
        from System.Windows.Forms import DialogResult, FolderBrowserDialog

        dlg = FolderBrowserDialog()
        dlg.Description = u""
        cur = u""
        try:
            cur = unicode(self._txt_carpeta.Text).strip()
        except Exception:
            cur = u""
        sel = u""
        if cur:
            try:
                if os.path.isdir(cur):
                    sel = cur
                else:
                    par = os.path.dirname(cur.rstrip(u"\\/"))
                    if par and os.path.isdir(par):
                        sel = par
            except Exception:
                sel = u""
        if sel:
            try:
                dlg.SelectedPath = sel
            except Exception:
                pass
        if dlg.ShowDialog() == DialogResult.OK:
            try:
                base = unicode(dlg.SelectedPath).strip()
                bn = os.path.basename(base.rstrip(u"\\/"))
                if _es_nombre_carpeta_entrega_estandar(bn):
                    self._txt_carpeta.Text = os.path.normpath(base)
                else:
                    suf = _nombre_carpeta_entrega_por_defecto()
                    self._txt_carpeta.Text = os.path.normpath(os.path.join(base, suf))
            except Exception:
                pass

    @staticmethod
    def _nullable_bool(wpf_nullable):
        """Convierte Nullable<bool> de WPF a bool de Python sin lanzar excepción."""
        try:
            if wpf_nullable is None:
                return False
            # HasValue / Value para Nullable<T>
            if hasattr(wpf_nullable, u"HasValue"):
                return bool(wpf_nullable.HasValue and wpf_nullable.Value)
            return unicode(wpf_nullable).strip().lower() == u"true"
        except Exception:
            return False

    def _row_is_selected(self, row):
        try:
            sel = row[u"Sel"]
            if self._nullable_bool(sel):
                return True
            if unicode(str(sel)).lower() == u"true":
                return True
        except Exception:
            pass
        return False

    def _on_grid_sel_preview_mouse_left_button_down(self, sender, e):
        """Un clic alterna Sel (evita el doble clic del DataGridCheckBoxColumn). Mayús+clic: rango.
        Si varias filas están seleccionadas en la cuadrícula, un clic en un checkbox aplica el mismo
        estado a todas ellas."""
        try:
            from System.Windows.Controls import DataGridCell, DataGridCheckBoxColumn
            from System.Windows.Input import Keyboard, ModifierKeys
            from System.Windows.Media import VisualTreeHelper
            from System import Boolean
        except Exception:
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

        g = self._grid
        if g is None:
            return
        try:
            idx = g.Items.IndexOf(drv)
        except Exception:
            idx = -1
        if idx < 0:
            return

        def _sel_get(rowv):
            try:
                if self._row_is_selected(rowv.Row):
                    return True
            except Exception:
                pass
            try:
                return self._nullable_bool(rowv[u"Sel"])
            except Exception:
                return False

        shift_down = False
        try:
            shift_down = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
        except Exception:
            pass

        anchor = getattr(self, u"_sel_anchor_idx", None)

        if shift_down and anchor is not None:
            try:
                cur_on = _sel_get(drv)
                new_val = not cur_on
                i0 = min(int(anchor), int(idx))
                i1 = max(int(anchor), int(idx))
                for j in range(i0, i1 + 1):
                    try:
                        rv = g.Items[j]
                        rv[u"Sel"] = Boolean(new_val)
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                e.Handled = True
            except Exception:
                pass
            try:
                self._refresh_estado()
            except Exception:
                pass
            return

        def _apply_checkbox_bulk():
            try:
                cur_on = _sel_get(drv)
                new_val = not cur_on
                self._sel_anchor_idx = idx

                indices = []
                seen = set()
                try:
                    coll = g.SelectedItems
                    if coll is not None:
                        for i in range(int(coll.Count)):
                            try:
                                it = coll[i]
                                j = g.Items.IndexOf(it)
                                if j >= 0 and j not in seen:
                                    seen.add(j)
                                    indices.append(j)
                            except Exception:
                                pass
                except Exception:
                    pass
                if idx not in seen:
                    seen.add(idx)
                    indices.append(idx)
                for j in indices:
                    try:
                        g.Items[j][u"Sel"] = Boolean(new_val)
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                self._refresh_estado()
            except Exception:
                pass

        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            g.Dispatcher.BeginInvoke(DispatcherPriority.Input, Action(_apply_checkbox_bulk))
        except Exception:
            _apply_checkbox_bulk()
        try:
            e.Handled = True
        except Exception:
            pass

    def _on_exportar(self, sender, args):
        try:
            self._exportar_impl()
        except Exception as _ex:
            try:
                _export_laminas_td_ok(
                    self._win,
                    u"Error inesperado:\n\n{0}".format(unicode(str(_ex))),
                )
            except Exception:
                pass

    def _exportar_impl(self):
        import os
        from Autodesk.Revit.DB import ElementId

        try:
            self._grid.CommitEdit()
        except Exception:
            pass

        do_pdf = self._nullable_bool(self._chk_pdf.IsChecked)
        do_dwg = self._nullable_bool(self._chk_dwg.IsChecked)
        do_listado = False
        if self._chk_listado_plan is not None:
            do_listado = self._nullable_bool(self._chk_listado_plan.IsChecked)

        if not do_pdf and not do_dwg and not do_listado:
            _export_laminas_td_ok(
                self._win,
                u"Marque al menos un formato: PDF, DWG o listado de planos.",
            )
            return

        entrega_root = u""
        try:
            entrega_root = unicode(self._txt_carpeta.Text).strip()
        except Exception:
            entrega_root = u""
        if not entrega_root:
            _export_laminas_td_ok(
                self._win,
                u"Indique la carpeta de entrega (ruta completa en el cuadro; use «Examinar…» si no la tiene).",
            )
            return
        try:
            entrega_root = os.path.normpath(entrega_root)
        except Exception:
            pass

        pdf_dir = os.path.join(entrega_root, u"PDF")
        dwg_dir = os.path.join(entrega_root, u"DWG")
        try:
            if do_pdf and not os.path.isdir(pdf_dir):
                os.makedirs(pdf_dir)
            if do_dwg and not os.path.isdir(dwg_dir):
                os.makedirs(dwg_dir)
            if do_listado and not do_pdf and not do_dwg:
                if not os.path.isdir(entrega_root):
                    os.makedirs(entrega_root)
        except Exception as ex:
            _export_laminas_td_ok(
                self._win,
                u"No se pudieron crear las carpetas:\n\n{}".format(unicode(str(ex))),
            )
            return

        selected_indices = []
        for i in range(self._table.Rows.Count):
            if self._row_is_selected(self._table.Rows[i]):
                selected_indices.append(i)

        n_sel = len(selected_indices)

        if do_listado and n_sel == 0:
            _export_laminas_td_ok(
                self._win,
                u"Para el listado Excel debe seleccionar al menos una lámina en la tabla.",
            )
            return

        if do_listado:
            if not os.path.isfile(_TEMPLATE_LISTADO_XLSX):
                _export_laminas_td_ok(
                    self._win,
                    u"No se encontró la plantilla de listado:\n{0}".format(_TEMPLATE_LISTADO_XLSX),
                )
                return
            if _listado_planos_core is None:
                _export_laminas_td_ok(
                    self._win,
                    u"No se pudo cargar el módulo de listado Excel (listado_planos_excel_core).",
                )
                return

        sheets_for_listado = []
        if do_listado:
            for i in selected_indices:
                row = self._table.Rows[i]
                try:
                    sid = int(row[u"IdInt"])
                except Exception:
                    continue
                try:
                    eid = ElementId(sid)
                except Exception:
                    try:
                        eid = ElementId(long(sid))
                    except Exception:
                        continue
                el = self._doc.GetElement(eid)
                if el is not None and isinstance(el, ViewSheet):
                    sheets_for_listado.append(el)
        steps_listado_on = bool(do_listado and n_sel > 0)
        # Tras DWG: barra de PDF con total = solo láminas PDF; el listado Excel es otra barra (1 paso), sin sumar al contador PDF.
        pbx = [None, False]  # [pb, pb_ok]

        def _phase_begin(title, count):
            _pbar_exit_safe(pbx[0], pbx[1])
            pbx[0] = None
            pbx[1] = False
            try:
                _c = int(count)
            except Exception:
                _c = 0
            if _c < 1:
                return
            try:
                pb = _export_laminas_pbar_start(title, _c)
                pbx[1] = bool(pb is not None)
                pbx[0] = pb
                if pbx[1] and pb is not None:
                    pb.__enter__()
            except Exception:
                pbx[0] = None
                pbx[1] = False

        def _phase_end():
            _pbar_exit_safe(pbx[0], bool(pbx[1]))
            pbx[0] = None
            pbx[1] = False

        def _phase_step(sn, total, base_title):
            if not pbx[1] or pbx[0] is None:
                return
            try:
                _export_laminas_pbar_step(pbx[0], sn[0], total, base_title)
            except Exception:
                pass
            sn[0] += 1

        if n_sel > 0:
            self._btn_export.IsEnabled = False
            self._chk_pdf.IsEnabled = False
            self._chk_dwg.IsEnabled = False
            if self._chk_listado_plan is not None:
                self._chk_listado_plan.IsEnabled = False

        n_pdf_ok = 0
        n_dwg_ok = 0
        n_listado_rows = 0
        listado_truncated = False
        listado_path = u""
        errores = []

        def _parse_row_sheet_job(i):
            """Devuelve (eid, custom) o None; si None, el llamador debe avanzar un paso de la fase actual."""
            row = self._table.Rows[i]
            try:
                sid = int(row[u"IdInt"])
            except Exception:
                errores.append(u"Fila sin Id válido.")
                return None
            try:
                eid = ElementId(sid)
            except Exception as _eid_ex:
                try:
                    eid = ElementId(long(sid))
                except Exception as _eid_ex2:
                    errores.append(
                        u"ElementId inválido (fila {0}, sid={1}): {2}".format(
                            i, sid, unicode(str(_eid_ex2))
                        )
                    )
                    return None
            try:
                custom = unicode(row[u"CustomName"]).strip()
            except Exception:
                custom = u""
            if not custom:
                custom = u"export"
            custom = self._sanitize_file_base(custom)
            return eid, custom

        try:
            if do_dwg and n_sel > 0:
                _phase_begin(
                    _export_laminas_pbar_initial_title(_EXP_LAM_PBAR_DWG, n_sel),
                    n_sel,
                )
                sn = [0]
                for i in selected_indices:
                    got = _parse_row_sheet_job(i)
                    if got is None:
                        _phase_step(
                            sn,
                            n_sel,
                            _EXP_LAM_PBAR_DWG,
                        )
                        continue
                    eid, custom = got
                    try:
                        ok_dwg = self._export_sheet_dwg(
                            self._doc,
                            dwg_dir,
                            eid,
                            custom,
                            _DWG_EXPORT_SETUP_NAME,
                        )
                        if ok_dwg:
                            n_dwg_ok += 1
                        else:
                            errores.append(
                                u"DWG — {0}: la exportación no generó archivo.".format(custom)
                            )
                    except Exception as ex:
                        errores.append(u"DWG — {0}: {1}".format(custom, unicode(str(ex))))
                    finally:
                        _phase_step(
                            sn,
                            n_sel,
                            _EXP_LAM_PBAR_DWG,
                        )
                _phase_end()

            if do_pdf and n_sel > 0:
                _phase_begin(
                    _export_laminas_pbar_initial_title(_EXP_LAM_PBAR_PDF, n_sel),
                    n_sel,
                )
                sn = [0]
                for i in selected_indices:
                    got = _parse_row_sheet_job(i)
                    if got is None:
                        _phase_step(
                            sn,
                            n_sel,
                            _EXP_LAM_PBAR_PDF,
                        )
                        continue
                    eid, custom = got
                    try:
                        result = self._export_sheet_pdf(self._doc, pdf_dir, eid, custom)
                        if result:
                            n_pdf_ok += 1
                        else:
                            errores.append(
                                u"PDF — {0}: la exportación devolvió False.".format(custom)
                            )
                    except Exception as ex:
                        errores.append(u"PDF — {0}: {1}".format(custom, unicode(str(ex))))
                    finally:
                        _phase_step(
                            sn,
                            n_sel,
                            _EXP_LAM_PBAR_PDF,
                        )
                _phase_end()

            if steps_listado_on:
                _phase_begin(
                    _export_laminas_pbar_initial_title(_EXP_LAM_PBAR_LISTADO, 1),
                    1,
                )
                sn = [0]
                if not sheets_for_listado:
                    errores.append(
                        u"Listado Excel: la selección no contiene láminas válidas para el listado."
                    )
                else:
                    try:
                        try:
                            proj = (self._doc.ProjectInformation.Name or u"Listado").strip()
                        except Exception:
                            proj = u"Listado"
                        proj_safe = re.sub(r'[<>:"/\\|?*]', u"_", proj).strip() or u"Listado"
                        listado_path = os.path.join(
                            entrega_root,
                            proj_safe + u"_ListadoPlanos.xlsx",
                        )
                        n_listado_rows, listado_truncated = _listado_planos_core.run_export_sheets(
                            __revit__,
                            _TEMPLATE_LISTADO_XLSX,
                            listado_path,
                            sheets_for_listado,
                        )
                    except Exception as ex:
                        errores.append(
                            u"Listado Excel: {0}".format(unicode(str(ex)))
                        )
                _phase_step(
                    sn,
                    1,
                    _EXP_LAM_PBAR_LISTADO,
                )
                _phase_end()
        finally:
            _pbar_exit_safe(pbx[0], bool(pbx[1]))
            pbx[0] = None
            pbx[1] = False
            if n_sel > 0:
                self._btn_export.IsEnabled = True
                self._chk_pdf.IsEnabled = True
                self._chk_dwg.IsEnabled = True
                if self._chk_listado_plan is not None:
                    self._chk_listado_plan.IsEnabled = True

        def _post_export_taskdialogs():
            if errores:
                try:
                    err_txt = u"\n".join(errores[:20])
                    if len(errores) > 20:
                        err_txt += u"\n…"
                    td_err = TaskDialog(_TASK_DLG_EXPORT_LAM_TITLE)
                    try:
                        td_err.TitleAutoPrefix = False
                    except Exception:
                        pass
                    td_err.MainInstruction = u"Se registraron errores durante la exportación."
                    td_err.MainContent = err_txt
                    td_err.CommonButtons = TaskDialogCommonButtons.Ok
                    td_err.DefaultButton = TaskDialogResult.Ok
                    td_err.Show()
                except Exception:
                    pass

            try:
                open_root = entrega_root
                try:
                    open_root = os.path.normpath(unicode(open_root).strip())
                except Exception:
                    open_root = entrega_root
                if open_root and os.path.isdir(open_root):
                    td_ask = TaskDialog(_TASK_DLG_EXPORT_LAM_TITLE)
                    try:
                        td_ask.TitleAutoPrefix = False
                    except Exception:
                        pass
                    td_ask.MainInstruction = (
                        u"¿Desea abrir la carpeta con todos los archivos generados?"
                    )
                    td_ask.MainContent = open_root
                    td_ask.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
                    td_ask.DefaultButton = TaskDialogResult.Yes
                    if td_ask.Show() == TaskDialogResult.Yes:
                        try:
                            os.startfile(open_root)
                        except Exception:
                            try:
                                import subprocess

                                subprocess.Popen([u"explorer", open_root])
                            except Exception:
                                pass
            except Exception:
                pass

        _run_revit_taskdialogs_above_wpf(self._win, _post_export_taskdialogs)
        self._refresh_estado()

    def show(self):
        try:
            from System.Windows.Interop import WindowInteropHelper

            _hw = revit_main_hwnd(__revit__.Application)
            if _hw:
                WindowInteropHelper(self._win).Owner = _hw
            position_wpf_window_top_left_at_active_view(self._win, uidoc, _hw)
        except Exception:
            pass
        try:
            sc = self._win.FindName("ExpLamRootScale")
            if sc is not None:
                sc.ScaleX = 0.0
                sc.ScaleY = 0.0
            self._win.Opacity = 0.0
        except Exception:
            pass
        try:
            System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, self._win)
        except Exception:
            pass
        self._win.Show()
        try:
            self._ensure_revit_cmd_blocker()
        except Exception:
            pass
        try:
            self._win.Activate()
        except Exception:
            pass


def main():
    existing = _get_active_tool_window()
    if existing is not None:
        ok = False
        try:
            from System.Windows import WindowState

            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
            existing.Show()
            existing.Activate()
            existing.Focus()
            ok = True
        except Exception:
            _clear_appdomain_window()
            existing = None
        if ok and existing is not None:
            _task_dialog_safe(
                _TASK_DLG_EXPORT_LAM_TITLE,
                u"La herramienta ya está en ejecución.",
                existing,
            )
            return

    w = ExportarLaminasWindow()
    w.show()


main()
