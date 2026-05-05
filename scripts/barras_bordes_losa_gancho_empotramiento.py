# -*- coding: utf-8 -*-
"""
Enfierrado pasada/shaft: formulario al abrir (misma línea de diseño UI que Crear Area Reinforcement RPS).

- Ancho del diálogo: misma lógica que fundación aislada (``_borde_losa_form_width_px`` ≈ 424 px con 3×110 mm + pads); altura SizeToContent.
- Información Armadura (esquema): cabecera = título + capas (derecha); contenido = Barras | diámetro alineado a la derecha con capas (mismos 110 mm + spinners).
- Título ventana y layout Arainco alineado con Enfierrado Vigas (tipografía, Combo 24 px, botón principal).
- Ventana no modal: selección host+caras vía botón + ExternalEvent; colocar enfierrado vía segundo evento.
- Capas y cantidad (set): incremental ▲/▼; capas muestra «N» + «Capa» o «Capas» (1 vs 2+); cantidad «N» + «Barras» y diámetro, ancho 110; rango 1–5; cantidad por defecto 2.

Revit 2024–2026 | pyRevit (IronPython 2.7). ElementId: .Value (2026+) o .IntegerValue.
"""

import os
import sys
import weakref
import clr
import System

# imp.load_source desde el .pushbutton no añade scripts/ al path; imports locales fallan sin esto.
_scripts_dir = os.path.dirname(os.path.abspath(__file__))
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
)
from Autodesk.Revit.DB.Structure import RebarBarType

from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_paths import get_logo_paths

_APPDOMAIN_WINDOW_KEY = "BIMTools.BordeLosaGanchoEmpotramiento.ActiveWindow"
_TOOL_TASK_DIALOG_TITLE = u"BIMTools — Refuerzo Borde Losa"


def _task_dialog_show(title, message, wpf_window=None):
    """
    TaskDialog de Revit queda detrás si el formulario WPF tiene Topmost=True.
    Se desactiva Topmost solo mientras dura el diálogo para que la alerta sea visible.
    """
    if wpf_window is not None:
        try:
            wpf_window.Topmost = False
        except Exception:
            pass
    try:
        TaskDialog.Show(title, message)
    finally:
        if wpf_window is not None:
            try:
                wpf_window.Topmost = True
            except Exception:
                pass


def element_id_to_int(element_id):
    """Entero estable para ElementId (2026+: Value; 2024–25: IntegerValue). Sin import extra."""
    if element_id is None:
        return None
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def _host_display_name(host):
    if host is None:
        return u""
    try:
        n = getattr(host, "Name", None)
        if n is not None:
            s = unicode(n).strip()
            if s:
                return s
    except Exception:
        pass
    try:
        return u"ID {0}".format(element_id_to_int(host.Id))
    except Exception:
        return u"(host)"


# Paleta Arainco / Revit 2024+: fondo #0A1A2F, acento cyan #5BC0DE, bordes #1A3A4D
# Combo: hover/foco #4C7383 (tenue); flecha #7AA3B8 — evita #5BC0DE en el campo (muy claro).
_CAPAS_OPCIONES_MIN = 1
_CAPAS_OPCIONES_MAX = 5
_BARRAS_SET_MIN = 1
_BARRAS_SET_MAX = 5
_WINDOW_OPEN_MS = 180
_WINDOW_CLOSE_MS = 180
_WINDOW_SLIDE_PX = 10.0
_MAX_BAR_LENGTH_DEFAULT_MM = 6000.0
_MAX_BAR_LENGTH_MIN_MM = 1000.0
_MAX_BAR_LENGTH_RECOMMENDED_MAX_MM = 12000.0
_LAP_MM_MIN = 100.0
_LAP_MM_MAX = 4000.0
_LAP_DEFAULT_MM = 860.0
_LAP_DETAIL_DEFAULT_FAMILY_NAME = u"EST_D_DEATIL ITEM_EMPALME"
_LAP_DETAIL_DEFAULT_TYPE_NAME = u"Empalme"
_LAP_DETAIL_ALT_FAMILY_NAMES = (
    u"EST_D_DEATIL ITEM_EMPALME",
    u"EST_D_DETAIL ITEM_EMPALME",
)

# Ancho: paralelo a ``enfierrado_fundacion_aislada._fundacion_aislada_form_width_px`` (3 columnas Capas|Barras|Ø tipo 110 px).
_BORDE_LOSA_INPUT_COLS = 3
_BORDE_LOSA_COMBO_W_PX = 110
_BORDE_LOSA_COL_GAP_PX = 10
_BORDE_LOSA_BLOCK_PAD_H_PX = 16
_BORDE_LOSA_GROUPBOX_PAD_H_PX = 16
_BORDE_LOSA_OUTER_PAD_H_PX = 28
_BORDE_LOSA_ARM_INFO_EXTRA_H_PX = 12
_BORDE_LOSA_TITLE_MIN_W_PX = 288


def _borde_losa_form_width_px():
    cols = max(1, int(_BORDE_LOSA_INPUT_COLS))
    c = int(_BORDE_LOSA_COMBO_W_PX)
    gaps = max(0, cols - 1) * int(_BORDE_LOSA_COL_GAP_PX)
    row_inner = cols * c + gaps + _BORDE_LOSA_BLOCK_PAD_H_PX
    w = (
        row_inner
        + _BORDE_LOSA_GROUPBOX_PAD_H_PX
        + _BORDE_LOSA_OUTER_PAD_H_PX
        + _BORDE_LOSA_ARM_INFO_EXTRA_H_PX
    )
    w = max(w, _BORDE_LOSA_TITLE_MIN_W_PX)
    return int((int(w) + 3) // 4 * 4)


def _normalize_capas_borde_losa_textbox(tb):
    """Acota capas al rango ``_CAPAS_OPCIONES_*`` (mismo criterio que el ComboBox previo)."""
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = unicode(_CAPAS_OPCIONES_MIN)
            return
        n = int(float(s.replace(u",", u".")))
        n = max(_CAPAS_OPCIONES_MIN, min(_CAPAS_OPCIONES_MAX, n))
        tb.Text = unicode(n)
    except Exception:
        tb.Text = unicode(_CAPAS_OPCIONES_MIN)


def _bump_capas_borde_losa_textbox(tb, delta):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if s:
            n = int(float(s.replace(u",", u".")))
        else:
            n = int(_CAPAS_OPCIONES_MIN)
    except Exception:
        n = int(_CAPAS_OPCIONES_MIN)
    n = max(
        _CAPAS_OPCIONES_MIN,
        min(_CAPAS_OPCIONES_MAX, n + int(delta)),
    )
    tb.Text = unicode(n)


_DEFAULT_BARRAS_SET_UI = 2


def _normalize_barras_set_borde_losa_textbox(tb):
    """Cantidad de barras del set (``_BARRAS_SET_MIN``–``_BARRAS_SET_MAX``)."""
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = unicode(_DEFAULT_BARRAS_SET_UI)
            return
        n = int(float(s.replace(u",", u".")))
        n = max(_BARRAS_SET_MIN, min(_BARRAS_SET_MAX, n))
        tb.Text = unicode(n)
    except Exception:
        tb.Text = unicode(_DEFAULT_BARRAS_SET_UI)


def _bump_barras_set_borde_losa_textbox(tb, delta):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if s:
            n = int(float(s.replace(u",", u".")))
        else:
            n = int(_DEFAULT_BARRAS_SET_UI)
    except Exception:
        n = int(_DEFAULT_BARRAS_SET_UI)
    n = max(
        _BARRAS_SET_MIN,
        min(_BARRAS_SET_MAX, n + int(delta)),
    )
    tb.Text = unicode(n)


# Misma línea visual que `enfierrado_vigas._ENFIERRADO_VIGAS_XAML` (tipografías, Combo 24 px, botones).
_ENFIERRADO_SHAFT_PASADA_XAML = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Refuerzo Borde Losa"
    Width="424"
    SizeToContent="Height"
    MinWidth="424" MaxWidth="424"
    WindowStartupLocation="Manual"
    Background="Transparent"
    AllowsTransparency="True"
    FontFamily="Segoe UI"
    WindowStyle="None"
    ResizeMode="NoResize"
    Topmost="True"
    UseLayoutRounding="True"
    >
  <Window.Resources>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
    <Style x:Key="ComboStretch" TargetType="ComboBox" BasedOn="{StaticResource Combo}">
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="HorizontalAlignment" Value="Stretch"/>
      <Setter Property="Height" Value="26"/>
      <Setter Property="MinHeight" Value="26"/>
      <Setter Property="MaxHeight" Value="26"/>
    </Style>
    <!-- Debe declararse antes de SpinnerSuffixTextBdLosa (StaticResource no admite referencia forward). -->
    <Style x:Key="SpinnerSuffixText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <!-- Mismo aspecto que el selector numérico Barras: borde redondeado + franja 18 px derecha (#0E1B32). -->
    <Style x:Key="SpinRepeatBtnBdLosa" TargetType="RepeatButton" BasedOn="{StaticResource SpinRepeatBtn}">
      <Setter Property="Width" Value="24"/>
    </Style>
    <Style x:Key="CantSpinnerTextBdLosa" TargetType="TextBox" BasedOn="{StaticResource CantSpinnerText}">
      <Setter Property="Padding" Value="6,0,4,0"/>
      <Setter Property="MinWidth" Value="28"/>
    </Style>
    <Style x:Key="SpinnerSuffixTextBdLosa" TargetType="TextBlock" BasedOn="{StaticResource SpinnerSuffixText}">
      <Setter Property="Margin" Value="4,0,10,0"/>
    </Style>
    <Style x:Key="ComboDiamSpinnerLook" TargetType="ComboBox" BasedOn="{StaticResource ComboStretch}">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid SnapsToDevicePixels="True">
              <Border x:Name="Border" CornerRadius="5" Background="#050E18"
                      BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True">
                <Grid TextElement.Foreground="{TemplateBinding Foreground}"
                      TextElement.FontWeight="{TemplateBinding FontWeight}">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="24"/>
                  </Grid.ColumnDefinitions>
                  <ContentPresenter x:Name="ContentSite" Grid.Column="0"
                                    Content="{TemplateBinding SelectionBoxItem}"
                                    ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                    Margin="8,0,6,0" VerticalAlignment="Center"
                                    HorizontalAlignment="Left" IsHitTestVisible="False"/>
                  <TextBox x:Name="PART_EditableTextBox"
                           Grid.Column="0" Visibility="Collapsed"
                           Background="Transparent" Foreground="{TemplateBinding Foreground}"
                           BorderThickness="0" Margin="8,0,6,0" VerticalAlignment="Center"
                           FontSize="{TemplateBinding FontSize}" FontFamily="{TemplateBinding FontFamily}"
                           CaretBrush="#7AA3B8" Padding="0" VerticalContentAlignment="Center"/>
                  <Border Grid.Column="1" Background="#11253D" BorderBrush="#1A3A4D"
                          BorderThickness="1,0,0,0" CornerRadius="0,5,5,0" ClipToBounds="True">
                    <TextBlock Text="&#9660;" FontSize="8" Foreground="#7AA3B8"
                               HorizontalAlignment="Center" VerticalAlignment="Center"
                               IsHitTestVisible="False"/>
                  </Border>
                  <ToggleButton Grid.Column="0" Grid.ColumnSpan="2"
                                IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                Focusable="False" Background="Transparent" BorderThickness="0"
                                HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>
                </Grid>
              </Border>
              <Popup x:Name="PART_Popup"
                     IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False"
                     PopupAnimation="Fade" Placement="Bottom"
                     PlacementTarget="{Binding ElementName=Border}">
                <Border Background="#050E18" BorderBrush="#1A3A4D" BorderThickness="1"
                        CornerRadius="5"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <ScrollViewer VerticalScrollBarVisibility="Auto">
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
                <Setter TargetName="Border" Property="Background" Value="#0B1728"/>
                <Setter TargetName="Border" Property="BorderBrush" Value="#4C7383"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocusWithin" Value="True">
                <Setter TargetName="Border" Property="Background" Value="#0B1728"/>
                <Setter TargetName="Border" Property="BorderBrush" Value="#4C7383"/>
                <Setter TargetName="Border" Property="BorderThickness" Value="2"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="Border" Property="Background" Value="#0B1728"/>
                <Setter TargetName="Border" Property="BorderBrush" Value="#4C7383"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="TbParamFill" TargetType="TextBox">
      <Setter Property="Background" Value="#050E18"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="BorderBrush" Value="#1A3A4D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Height" Value="26"/>
      <Setter Property="Padding" Value="8,0,10,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="CaretBrush" Value="#7AA3B8"/>
    </Style>
    <Style x:Key="TbParamNoBorder" TargetType="TextBox" BasedOn="{StaticResource TbParamFill}">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Padding" Value="8,0,6,0"/>
      <Setter Property="Height" Value="24"/>
    </Style>
    <Style x:Key="ComboDiamSpinnerLookFull" TargetType="ComboBox" BasedOn="{StaticResource ComboDiamSpinnerLook}">
      <Setter Property="HorizontalAlignment" Value="Stretch"/>
      <Setter Property="Width" Value="NaN"/>
      <Setter Property="MinWidth" Value="0"/>
      <Setter Property="MaxWidth" Value="99999"/>
    </Style>
    <!-- Aviso troceo: caja cyan (#5BC0DE) y texto oscuro (mockup Refuerzo Borde Losa) -->
    <Style x:Key="BorderAlertaTroceo" TargetType="Border">
      <Setter Property="Background" Value="#5BC0DE"/>
      <Setter Property="BorderBrush" Value="#49B1D0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="6"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
    </Style>
  </Window.Resources>
  <!-- Ancho alineado a fundación aislada (424 px vía _borde_losa_form_width_px en código) -->
  <Border CornerRadius="10" Background="#0A1A2F" Padding="12"
          BorderBrush="#1A3A4D" BorderThickness="1"
          HorizontalAlignment="Stretch" ClipToBounds="True">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="16" ShadowDepth="0" Opacity="0.35"/>
    </Border.Effect>

    <Grid HorizontalAlignment="Stretch">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border x:Name="TitleBar" Grid.Row="0" Background="#0E1B32" CornerRadius="6" Padding="10,8" Margin="0,0,0,10"
            BorderBrush="#21465C" BorderThickness="1" HorizontalAlignment="Stretch">
      <Grid HorizontalAlignment="Stretch">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Image x:Name="ImgLogo" Width="40" Height="40" Grid.Column="0"
               Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center"/>
        <TextBlock Grid.Column="1" Text="Refuerzo Borde Losa"
                   FontSize="15" FontWeight="SemiBold" Foreground="#E8F4F8"
                   VerticalAlignment="Center" HorizontalAlignment="Left" Margin="0,0,8,0"
                   TextWrapping="Wrap"/>
        <Button x:Name="BtnClose"
                Grid.Column="2"
                Style="{StaticResource BtnCloseX_MinimalNoBg}"
                VerticalAlignment="Center"
                HorizontalAlignment="Right"
                ToolTip="Cerrar"/>
      </Grid>
    </Border>

    <StackPanel Grid.Row="1" Margin="0,0,0,0" HorizontalAlignment="Stretch">
      <Button x:Name="BtnSeleccionar" Content="Seleccionar host y caras"
              Style="{StaticResource BtnSelectOutline}"
              HorizontalAlignment="Stretch" Margin="0,0,0,6">
        <Button.ToolTip>
          <ToolTip Background="#0E1B32" BorderBrush="#1A3A4D" Foreground="#C8E4EF" MaxWidth="340" Padding="10,8">
            <TextBlock Text="Primero el elemento host (p. ej. losa). Después las caras del hueco en ese mismo elemento."
                       TextWrapping="Wrap" FontSize="11" Foreground="#95B8CC"/>
          </ToolTip>
        </Button.ToolTip>
      </Button>
      <TextBlock x:Name="TxtSeleccionLine" Text="" Foreground="#7AD7A8" FontSize="11" Margin="0,0,0,10"
                 TextWrapping="Wrap" LineHeight="15"/>

    <!-- Aire inferior ~18 px antes del pie (misma sensación que card + fila botón en fundación aislada) -->
    <GroupBox Style="{StaticResource GbParams}" Margin="0,0,0,18" HorizontalAlignment="Stretch">
      <GroupBox.Header>
        <TextBlock Text="Información armadura" FontWeight="SemiBold"
                   Foreground="#E8F4F8" FontSize="11" VerticalAlignment="Center"/>
      </GroupBox.Header>
      <Grid Margin="0,4,0,0" HorizontalAlignment="Stretch">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*" MinWidth="88"/>
          <ColumnDefinition Width="12"/>
          <ColumnDefinition Width="*" MinWidth="118"/>
          <ColumnDefinition Width="12"/>
          <ColumnDefinition Width="*" MinWidth="100"/>
        </Grid.ColumnDefinitions>
        <Border Grid.Column="0" MinWidth="88" Height="26"
                HorizontalAlignment="Stretch" VerticalAlignment="Center"
                CornerRadius="4" Background="#050E18"
                BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True">
          <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="24"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="TxtNumCapasBordeLosa" Grid.Column="0" Style="{StaticResource CantSpinnerTextBdLosa}"
                     Width="28" Text="1" TextAlignment="Center" VerticalContentAlignment="Center"/>
            <TextBlock x:Name="TbCapasSuffixBordeLosa" Grid.Column="1" Text="Capa"
                       Style="{StaticResource SpinnerSuffixTextBdLosa}"/>
            <Border Grid.Column="3" Background="#0E1B32" BorderBrush="#1A3A4D"
                    BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <RepeatButton x:Name="BtnNumCapasBordeLosaUp" Grid.Row="0"
                              Style="{StaticResource SpinRepeatBtnBdLosa}" Content="▲" ToolTip="Más capas"/>
                <RepeatButton x:Name="BtnNumCapasBordeLosaDown" Grid.Row="1"
                              Style="{StaticResource SpinRepeatBtnBdLosa}" Content="▼" ToolTip="Menos capas"/>
              </Grid>
            </Border>
          </Grid>
        </Border>
        <Border Grid.Column="2" MinWidth="112" Height="26"
                HorizontalAlignment="Stretch" VerticalAlignment="Center"
                CornerRadius="4" Background="#050E18"
                BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True">
          <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="24"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="TxtCantidadBarrasSet" Grid.Column="0" Style="{StaticResource CantSpinnerTextBdLosa}"
                     Width="28" Text="2" TextAlignment="Center" VerticalContentAlignment="Center"/>
            <TextBlock Grid.Column="1" Text="Barras" Style="{StaticResource SpinnerSuffixTextBdLosa}"/>
            <Border Grid.Column="3" Background="#0E1B32" BorderBrush="#1A3A4D"
                    BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <RepeatButton x:Name="BtnCantidadBarrasSetUp" Grid.Row="0"
                              Style="{StaticResource SpinRepeatBtnBdLosa}" Content="▲" ToolTip="Más cantidad"/>
                <RepeatButton x:Name="BtnCantidadBarrasSetDown" Grid.Row="1"
                              Style="{StaticResource SpinRepeatBtnBdLosa}" Content="▼" ToolTip="Menos cantidad"/>
              </Grid>
            </Border>
          </Grid>
        </Border>
        <ComboBox x:Name="CmbBarType" Grid.Column="4" Style="{StaticResource ComboDiamSpinnerLookFull}"
                  VerticalAlignment="Center">
          <ComboBox.ItemContainerStyle>
            <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
          </ComboBox.ItemContainerStyle>
        </ComboBox>
      </Grid>
    </GroupBox>

      <StackPanel x:Name="PnlMaxBarLength" Visibility="Collapsed" Margin="0,0,0,18">
        <GroupBox Style="{StaticResource GbParams}" Margin="0,0,0,8" HorizontalAlignment="Stretch">
          <GroupBox.Header>
            <TextBlock Text="Parámetros de longitud" FontWeight="SemiBold"
                       Foreground="#E8F4F8" FontSize="11" VerticalAlignment="Center"/>
          </GroupBox.Header>
          <Grid Margin="0,4,0,0" HorizontalAlignment="Stretch">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Row="0" Grid.Column="0" Text="Largo máx. tramo (mm)" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <TextBlock Grid.Row="0" Grid.Column="2" Text="Empalme / traslape (mm)" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <Border Grid.Row="1" Grid.Column="0" CornerRadius="4" Background="#050E18" BorderBrush="#1A3A4D"
                    BorderThickness="1" Height="26" SnapsToDevicePixels="True">
              <TextBox x:Name="TxtMaxBarLength" Style="{StaticResource TbParamNoBorder}"
                       ToolTip="Use un valor entre 1000 y 12000 mm"/>
            </Border>
            <Border Grid.Row="1" Grid.Column="2" CornerRadius="4" Background="#050E18" BorderBrush="#1A3A4D"
                    BorderThickness="1" Height="26" SnapsToDevicePixels="True">
              <TextBox x:Name="TxtLapMm" Style="{StaticResource TbParamNoBorder}"
                       ToolTip="Según tabla por Ø (Información armadura); puede ajustarse manualmente."/>
            </Border>
          </Grid>
        </GroupBox>
        <Border x:Name="BorderTroceoAviso" Visibility="Collapsed" Style="{StaticResource BorderAlertaTroceo}" Margin="0,0,0,0">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="&#x26A0;" FontSize="18" Foreground="#0A1A2F"
                       VerticalAlignment="Top" Margin="0,0,8,0" LineHeight="20"/>
            <TextBlock x:Name="TxtTroceoAviso" Grid.Column="1" Foreground="#0A1A2F" FontSize="11"
                       TextWrapping="Wrap" LineHeight="16"/>
          </Grid>
        </Border>
      </StackPanel>

      <!-- Estado opcional (como BorderPropagacion); el hueco card→botón lo da el GroupBox Información (mb 18) o Pnl (mb 18) -->
      <TextBlock x:Name="TxtEstado" Text="" Foreground="#5BC0DE" FontSize="11"
                 Margin="0,0,0,6" TextWrapping="Wrap" Visibility="Collapsed"/>
      <Button x:Name="BtnGenerar" Content="Colocar armaduras"
              Style="{StaticResource BtnPrimary}"
              HorizontalAlignment="Stretch"/>
    </StackPanel>

    </Grid>
  </Border>
</Window>
"""


def _rebar_bar_type_display_name(bt):
    if bt is None:
        return u""
    try:
        n = bt.Name
        if n is not None and unicode(n).strip():
            return unicode(n).strip()
    except Exception:
        pass
    for bip_name in ("SYMBOL_NAME_PARAM", "ALL_MODEL_TYPE_NAME", "ELEM_TYPE_PARAM"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = bt.get_Parameter(bip)
            if p is not None and p.HasValue:
                s = p.AsString()
                if s is not None and unicode(s).strip():
                    return unicode(s).strip()
        except Exception:
            continue
    try:
        d_mm = int(round(float(bt.BarNominalDiameter) * 304.8))
        if d_mm > 0:
            return u"\u00f8{0} mm".format(d_mm)
    except Exception:
        pass
    try:
        return u"ID {0}".format(element_id_to_int(bt.Id))
    except Exception:
        return u"(tipo sin nombre)"


def _rebar_nominal_diameter_mm(bt):
    try:
        d_mm = int(round(float(bt.BarNominalDiameter) * 304.8))
        return d_mm if d_mm > 0 else None
    except Exception:
        return None


def _normalize_lap_tb_borde(tb):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = unicode(int(_LAP_DEFAULT_MM))
            return
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        tb.Text = unicode(int(_LAP_DEFAULT_MM))
        return
    n = max(int(_LAP_MM_MIN), min(int(_LAP_MM_MAX), n))
    tb.Text = unicode(int(n))


def _read_lap_tb_borde(tb):
    if tb is None:
        return float(_LAP_DEFAULT_MM)
    try:
        s = unicode(tb.Text).strip()
        if not s:
            return float(_LAP_DEFAULT_MM)
        n = float(s.replace(u",", u"."))
    except Exception:
        return float(_LAP_DEFAULT_MM)
    n = max(_LAP_MM_MIN, min(_LAP_MM_MAX, n))
    return n


def _borde_losa_traslape_mm(d_bar_mm, tlap_tb):
    """Traslape (mm) según tabla por Ø nominal de la barra (combo principal); respaldo ``TxtLapMm``."""
    try:
        if d_bar_mm is not None and float(d_bar_mm) > 1e-6:
            from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm

            v = traslape_mm_from_nominal_diameter_mm(float(d_bar_mm), None)
            if v is not None:
                return float(v)
    except Exception:
        pass
    return _read_lap_tb_borde(tlap_tb)


def _is_digits_only_text(s):
    if s is None:
        return False
    t = unicode(s)
    if not t:
        return False
    for ch in t:
        if ch < u"0" or ch > u"9":
            return False
    return True


def _rebar_bar_type_combo_label(bt):
    """Mismo criterio visual que Crear Area Reinforcement RPS: øN mm (pies → mm)."""
    if bt is None:
        return u""
    d_mm = _rebar_nominal_diameter_mm(bt)
    if d_mm is not None:
        return u"\u00f8{0} mm".format(d_mm)
    try:
        return u"Barra (ID{0})".format(element_id_to_int(bt.Id))
    except Exception:
        return u"\u2014"


def _build_bar_type_entries(doc):
    tipos = list(FilteredElementCollector(doc).OfClass(RebarBarType))
    if not tipos:
        return None, u"No hay RebarBarType en el proyecto."
    # La fila «ø12 mm» (automático) cubre el tipo más cercano a 12 mm; no duplicarlo en la lista.
    auto_id = None
    try:
        from enfierrado_shaft_hashtag import resolver_bar_type_por_diametro_mm

        bt_auto, _, _ = resolver_bar_type_por_diametro_mm(doc, 12.0)
        if bt_auto is not None:
            auto_id = element_id_to_int(bt_auto.Id)
    except Exception:
        pass
    tipos = [
        bt
        for bt in tipos
        if auto_id is None
        or element_id_to_int(bt.Id) != auto_id
    ]
    base_labels = [_rebar_bar_type_combo_label(bt) for bt in tipos]
    key_counts = {}
    for lb in base_labels:
        k = lb.lower()
        key_counts[k] = key_counts.get(k, 0) + 1
    typed = []
    for bt, lbl in zip(tipos, base_labels):
        if key_counts.get(lbl.lower(), 0) > 1:
            lbl = u"{0}  [Id {1}]".format(lbl, element_id_to_int(bt.Id))
        typed.append((bt, lbl))

    def _sort_key(entry):
        """Orden creciente por BarNominalDiameter (pies), igual que _get_rebar_bar_types (Area RPS)."""
        bt, _lbl = entry
        try:
            d_ft = float(bt.BarNominalDiameter)
        except Exception:
            d_ft = 0.0
        return (d_ft, element_id_to_int(bt.Id) or 0)

    typed.sort(key=_sort_key)
    # Misma lógica de orden que Area RPS, incluyendo la fila automática «12 mm nominal»
    # en su puesto (no al principio del todo).
    auto_target_mm = 12.0
    try:
        auto_ft = float(auto_target_mm) / 304.8
    except Exception:
        auto_ft = 12.0 / 304.8
    auto_lbl = u"\u00f812 mm"
    rows = []
    rows.append((auto_ft, 0, 0, (None, auto_lbl)))
    for bt, lbl in typed:
        try:
            d_ft = float(bt.BarNominalDiameter)
        except Exception:
            d_ft = 0.0
        oid = element_id_to_int(bt.Id) or 0
        rows.append((d_ft, 1, oid, (bt, lbl)))
    rows.sort(key=lambda r: (r[0], r[1], r[2]))
    entries = [r[3] for r in rows]
    return entries, None


def _find_fixed_lap_detail_symbol_id(doc):
    if doc is None:
        return None, u"No hay documento activo."
    fam_target = unicode(_LAP_DETAIL_DEFAULT_FAMILY_NAME or u"").strip().lower()
    fam_alt_targets = set()
    try:
        for nm in _LAP_DETAIL_ALT_FAMILY_NAMES:
            t = unicode(nm or u"").strip().lower()
            if t:
                fam_alt_targets.add(t)
    except Exception:
        pass
    if fam_target:
        fam_alt_targets.add(fam_target)
    typ_target = unicode(_LAP_DETAIL_DEFAULT_TYPE_NAME or u"").strip().lower()

    def _norm_name(s):
        try:
            t = unicode(s or u"")
        except Exception:
            t = u""
        try:
            t = t.replace(u"\u00A0", u" ")
        except Exception:
            pass
        t = u" ".join([p for p in t.strip().lower().split() if p])
        return t

    try:
        syms = list(
            FilteredElementCollector(doc)
            .OfClass(FamilySymbol)
            .OfCategory(BuiltInCategory.OST_DetailComponents)
        )
    except Exception:
        syms = []
    if not syms:
        return None, u"No hay Detail Components en el proyecto."
    fam_alt_targets = set([_norm_name(x) for x in fam_alt_targets if _norm_name(x)])
    fam_target = _norm_name(fam_target)
    typ_target = _norm_name(typ_target)
    for sym in syms:
        if sym is None:
            continue
        fam = u""
        typ = u""
        try:
            fam = _norm_name(getattr(sym, "FamilyName", None))
        except Exception:
            fam = u""
        if not fam:
            try:
                if sym.Family is not None:
                    fam = _norm_name(sym.Family.Name)
            except Exception:
                fam = u""
        try:
            typ = _norm_name(getattr(sym, "Name", None))
        except Exception:
            typ = u""
        fam_ok = (fam in fam_alt_targets) if fam_alt_targets else (fam == fam_target)
        if fam_ok and typ == typ_target:
            try:
                return sym.Id, None
            except Exception:
                break
    for sym in syms:
        if sym is None:
            continue
        try:
            fam = _norm_name(getattr(sym, "FamilyName", None))
        except Exception:
            fam = u""
        if not fam:
            try:
                if sym.Family is not None:
                    fam = _norm_name(sym.Family.Name)
            except Exception:
                fam = u""
        fam_ok = (fam in fam_alt_targets) if fam_alt_targets else (fam == fam_target)
        if fam_ok:
            try:
                return sym.Id, (
                    u"Detail Component fijo: no se encontró tipo exacto '{0}', se usó otro tipo de la familia '{1}'."
                    .format(_LAP_DETAIL_DEFAULT_TYPE_NAME, _LAP_DETAIL_DEFAULT_FAMILY_NAME)
                )
            except Exception:
                pass
    return None, (
        u"No se encontró Detail Component fijo '{0} : {1}'."
        .format(_LAP_DETAIL_DEFAULT_FAMILY_NAME, _LAP_DETAIL_DEFAULT_TYPE_NAME)
    )


class SeleccionarHostCarasShaftHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        from seleccion_caras_elemento import seleccionar_host_y_caras

        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win._set_estado(u"No hay documento activo.")
            return
        doc = uidoc.Document
        try:
            host, refs = seleccionar_host_y_caras(
                uidoc, doc, guardar_en_modulo=True, mostrar_errores=False
            )
            if host is not None and refs:
                win._host = host
                win._refs = refs
                win._max_length_confirmed = False
                hname = _host_display_name(host)
                win._set_seleccion_line(
                    u"\u2713 {0} \u00b7 {1} cara(s) seleccionada(s)".format(hname, len(refs))
                )
                win._set_estado(u"")
                win._actualizar_info_host_caras()
                win._refresh_max_length_ui_from_current_selection()
            else:
                win._set_seleccion_line(u"")
                win._set_estado(u"Selección cancelada o sin caras. La selección anterior no se modificó.")
        except Exception as ex:
            win._set_estado(u"Error en selección: {0}".format(ex))
            try:
                _task_dialog_show(
                    _TOOL_TASK_DIALOG_TITLE,
                    u"Error en selección:\n{0}".format(ex),
                    win._win,
                )
            except Exception:
                pass
        finally:
            try:
                win._show_with_fade()
            except Exception:
                pass

    def GetName(self):
        return u"SeleccionarHostCarasShaftPasada"


class GenerarEnfierradoShaftHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref
        self.n_capas = 1
        self.n_barras_set = 2
        self.forced_bar_type_id = None
        self.max_bar_length_mm = float(_MAX_BAR_LENGTH_DEFAULT_MM)
        self.lap_length_mm = None
        self.place_lap_details = True
        self.lap_detail_symbol_id = None

    def Execute(self, uiapp):
        from enfierrado_shaft_hashtag import crear_enfierrado_bordes_losa_gancho_y_empotramiento

        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            _task_dialog_show(
                _TOOL_TASK_DIALOG_TITLE,
                u"No hay documento activo.",
                win._win,
            )
            return
        doc = uidoc.Document
        host = win._host
        refs = win._refs
        titulo = _TOOL_TASK_DIALOG_TITLE

        if host is None or not refs:
            _task_dialog_show(
                titulo,
                u"Seleccione primero el host y al menos una cara con el botón correspondiente.",
                win._win,
            )
            return

        n_capas = max(1, int(self.n_capas))
        n_barras_set = max(_BARRAS_SET_MIN, min(_BARRAS_SET_MAX, int(self.n_barras_set or 2)))
        if n_capas == 2:
            tag_family_name = u"EST_A_STRUCTURAL REBAR TAG_DOUBLE QUANTITY"
        elif n_capas == 3:
            tag_family_name = u"EST_A_STRUCTURAL REBAR TAG_TRIPLE QUANTITY"
        elif n_capas == 4:
            tag_family_name = u"EST_A_STRUCTURAL REBAR TAG_QUADRUPLE QUANTITY"
        elif n_capas == 5:
            tag_family_name = u"EST_A_STRUCTURAL REBAR TAG_QUINTUPLE QUANTITY"
        else:
            tag_family_name = u"EST_A_STRUCTURAL REBAR TAG"
        bar_type_id = self.forced_bar_type_id
        if bar_type_id is not None and bar_type_id == ElementId.InvalidElementId:
            bar_type_id = None

        try:
            creados, n_tags, _rebar_ids, avisos_rb, err_rb = crear_enfierrado_bordes_losa_gancho_y_empotramiento(
                doc,
                host,
                refs,
                n_capas=n_capas,
                n_barras_set=n_barras_set,
                forced_bar_type_id=bar_type_id,
                max_bar_length_mm=float(self.max_bar_length_mm or _MAX_BAR_LENGTH_DEFAULT_MM),
                lap_length_mm=self.lap_length_mm,
                tag_view=uidoc.ActiveView,
                tag_family_name=tag_family_name,
                place_lap_details=bool(self.place_lap_details),
                lap_detail_view=uidoc.ActiveView,
                lap_detail_symbol_id=self.lap_detail_symbol_id,
            )
        except Exception as ex:
            _task_dialog_show(titulo, u"Error:\n\n{0}".format(ex), win._win)
            try:
                win._set_estado(u"")
            except Exception:
                pass
            return

        if err_rb:
            err_txt = unicode(err_rb).strip()
            if len(err_txt) > 400:
                err_txt = err_txt[:397] + u"…"
            _task_dialog_show(titulo, u"Incidencias:\n\n{0}".format(err_txt), win._win)
            try:
                win._set_estado(u"")
            except Exception:
                pass
            return

        # Mostrar advertencias de etiquetado y cotas si alguna no se pudo crear.
        try:
            _kw_warn = (
                u"tag", u"etiqueta", u"familia", u"shape", u"fallo al crear",
                u"cota", u"cotas", u"empotramiento", u"empalme", u"traslapo",
                u"planta", u"omitió",
            )
            tag_warns = [
                unicode(av).strip()
                for av in (avisos_rb or [])
                if any(kw in unicode(av).lower() for kw in _kw_warn)
            ]
            n_tags_int = int(n_tags or 0)
            if tag_warns:
                warn_txt = u" | ".join(tag_warns[:3])
                if len(warn_txt) > 220:
                    warn_txt = warn_txt[:217] + u"…"
                win._set_estado(
                    u"Tags: {0}. {1}".format(n_tags_int, warn_txt)
                )
            elif n_tags_int == 0 and int(creados or 0) > 0:
                win._set_estado(
                    u"Barras creadas ({0}). Sin etiquetas: verifique que la familia '{1}' esté cargada.".format(
                        int(creados), tag_family_name
                    )
                )
            else:
                win._set_estado(u"")
        except Exception:
            try:
                win._set_estado(u"")
            except Exception:
                pass
        try:
            win._max_length_confirmed = True
        except Exception:
            pass
        # Evita duplicados en ejecuciones consecutivas: tras crear barras con éxito,
        # forzamos nueva selección de host/caras.
        try:
            if int(creados or 0) > 0:
                win._limpiar_seleccion_host_caras()
        except Exception:
            pass

    def GetName(self):
        return u"GenerarEnfierradoShaftPasada"


class EnfierradoShaftPasadaWindow(object):
    def __init__(self, revit):
        self._revit = revit
        self._document = None
        self._host = None
        self._refs = None
        self._entries = []
        self._bar_sel_hooked = False
        self._is_closing_with_fade = False
        self._base_top = None
        self._default_max_bar_length_mm = float(_MAX_BAR_LENGTH_DEFAULT_MM)
        self._max_length_confirmed = False

        from System.Windows import RoutedEventHandler
        from System.Windows.Input import ApplicationCommands, CommandBinding, Key, KeyBinding, ModifierKeys
        from System.Windows.Markup import XamlReader

        self._win = XamlReader.Parse(_ENFIERRADO_SHAFT_PASADA_XAML)
        try:
            wpx = float(_borde_losa_form_width_px())
            self._win.Width = wpx
            self._win.MinWidth = wpx
            self._win.MaxWidth = wpx
        except Exception:
            pass

        self._seleccion_handler = SeleccionarHostCarasShaftHandler(weakref.ref(self))
        self._seleccion_event = ExternalEvent.Create(self._seleccion_handler)
        self._generar_handler = GenerarEnfierradoShaftHandler(weakref.ref(self))
        self._generar_event = ExternalEvent.Create(self._generar_handler)

        self._setup_ui()
        self._wire_commands(RoutedEventHandler, ApplicationCommands, CommandBinding, KeyBinding, Key, ModifierKeys)
        self._wire_lifecycle_handlers()

    def _wire_lifecycle_handlers(self):
        try:
            from System.Windows import RoutedEventHandler

            def _on_closed(sender, args):
                try:
                    _clear_active_window()
                except Exception:
                    pass

            self._win.Closed += RoutedEventHandler(_on_closed)
        except Exception:
            pass

    def _setup_ui(self):
        from System.IO import FileAccess, FileMode, FileStream
        from System.Windows import RoutedEventHandler, DataObject, DataObjectPastingEventHandler
        from System.Windows.Input import TextCompositionEventHandler
        from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage
        logo_loaded = False
        try:
            img = self._win.FindName("ImgLogo")
            if img is not None:
                for logo_path in get_logo_paths():
                    if os.path.isfile(logo_path):
                        stream = None
                        try:
                            stream = FileStream(
                                logo_path,
                                FileMode.Open,
                                FileAccess.Read,
                            )
                            bmp = BitmapImage()
                            bmp.BeginInit()
                            bmp.StreamSource = stream
                            bmp.CacheOption = BitmapCacheOption.OnLoad
                            bmp.EndInit()
                            bmp.Freeze()
                            img.Source = bmp
                            logo_loaded = True
                        finally:
                            if stream is not None:
                                try:
                                    stream.Dispose()
                                except Exception:
                                    try:
                                        stream.Close()
                                    except Exception:
                                        pass
                        break
        except Exception:
            pass
        if not logo_loaded:
            try:
                from pyrevit import script

                script.get_logger().warn(
                    u"[barras_bordes_losa_gancho_empotramiento] Ningún logo encontrado. Coloque "
                    u"empresa_logo.png, logo_empresa.png o logo.png en la carpeta del botón: "
                    + _ENFIERRADO_PASADA_PUSHBUTTON
                )
            except Exception:
                pass

        btn_sel = self._win.FindName("BtnSeleccionar")
        if btn_sel is not None:
            btn_sel.Click += RoutedEventHandler(self._on_seleccionar)
        btn_gen = self._win.FindName("BtnGenerar")
        if btn_gen is not None:
            btn_gen.Click += RoutedEventHandler(self._on_generar)
        txt_max_len = self._win.FindName("TxtMaxBarLength")
        if txt_max_len is not None:
            try:
                txt_max_len.Text = unicode(int(round(float(self._default_max_bar_length_mm))))
            except Exception:
                txt_max_len.Text = u"6000"

            def _preview_text_input(sender, e):
                try:
                    e.Handled = not _is_digits_only_text(e.Text)
                except Exception:
                    pass

            def _on_pasting(sender, e):
                try:
                    txt = None
                    if e.SourceDataObject is not None and e.SourceDataObject.GetDataPresent("Text"):
                        txt = e.SourceDataObject.GetData("Text")
                    if not _is_digits_only_text(txt):
                        e.CancelCommand()
                except Exception:
                    try:
                        e.CancelCommand()
                    except Exception:
                        pass

            try:
                txt_max_len.PreviewTextInput += TextCompositionEventHandler(_preview_text_input)
            except Exception:
                pass
            try:
                DataObject.AddPastingHandler(txt_max_len, DataObjectPastingEventHandler(_on_pasting))
            except Exception:
                pass

        txt_lap = self._win.FindName("TxtLapMm")
        if txt_lap is not None:
            try:
                txt_lap.Text = unicode(int(round(float(_LAP_DEFAULT_MM))))
            except Exception:
                txt_lap.Text = unicode(int(_LAP_DEFAULT_MM))

            def _preview_lap_input(sender, e):
                try:
                    e.Handled = not _is_digits_only_text(e.Text)
                except Exception:
                    pass

            def _on_lap_pasting(sender, e):
                try:
                    txt = None
                    if e.SourceDataObject is not None and e.SourceDataObject.GetDataPresent("Text"):
                        txt = e.SourceDataObject.GetData("Text")
                    if not _is_digits_only_text(txt):
                        e.CancelCommand()
                except Exception:
                    try:
                        e.CancelCommand()
                    except Exception:
                        pass

            try:
                from System.Windows import RoutedEventHandler

                txt_lap.LostFocus += RoutedEventHandler(
                    lambda s, a, tbx=txt_lap: _normalize_lap_tb_borde(tbx)
                )
            except Exception:
                pass
            try:
                txt_lap.PreviewTextInput += TextCompositionEventHandler(_preview_lap_input)
            except Exception:
                pass
            try:
                DataObject.AddPastingHandler(txt_lap, DataObjectPastingEventHandler(_on_lap_pasting))
            except Exception:
                pass

        try:
            from System.Windows.Automation import AutomationProperties

            if btn_sel is not None:
                AutomationProperties.SetName(btn_sel, u"Seleccionar host y caras")
            txt_sel = self._win.FindName("TxtSeleccionLine")
            if txt_sel is not None:
                AutomationProperties.SetName(txt_sel, u"Resumen de selección de host y caras")
            if btn_gen is not None:
                AutomationProperties.SetName(btn_gen, u"Colocar enfierrado")
            tb_capas = self._win.FindName("TxtNumCapasBordeLosa")
            if tb_capas is not None:
                AutomationProperties.SetName(
                    tb_capas,
                    u"Número de capas (1 a 5); «Capa» o «Capas» junto al valor es solo informativo",
                )
            tb_cant = self._win.FindName("TxtCantidadBarrasSet")
            if tb_cant is not None:
                AutomationProperties.SetName(
                    tb_cant, u"Número de barras del set (1 a 5); «Barras» es solo informativo junto al valor"
                )
            cmb_bar = self._win.FindName("CmbBarType")
            if cmb_bar is not None:
                AutomationProperties.SetName(cmb_bar, u"Diámetro de barra (ø mm)")
            txt_lap_a11y = self._win.FindName("TxtLapMm")
            if txt_lap_a11y is not None:
                AutomationProperties.SetName(txt_lap_a11y, u"Empalme o traslape en milímetros")
        except Exception:
            pass

        self._wire_capas_spinner_once()
        self._wire_barras_set_spinner_once()

        # Sin chrome del sistema (WindowStyle=None), la ventana debe ser movible
        # desde una zona propia. El botón Cerrar no debe disparar el arrastre.
        try:
            from System.Windows.Input import MouseButtonEventHandler

            btn_close = self._win.FindName("BtnClose")
            title_bar = self._win.FindName("TitleBar")

            if title_bar is not None:
                def _on_titlebar_down(sender, e):
                    try:
                        self._win.DragMove()
                    except Exception:
                        pass

                title_bar.MouseLeftButtonDown += MouseButtonEventHandler(
                    _on_titlebar_down
                )

            if btn_close is not None:
                def _on_close_click(sender, e):
                    try:
                        self._close_with_fade()
                    except Exception:
                        pass

                btn_close.Click += RoutedEventHandler(_on_close_click)

                def _on_close_down(sender, e):
                    try:
                        e.Handled = True
                    except Exception:
                        pass

                btn_close.MouseLeftButtonDown += MouseButtonEventHandler(
                    _on_close_down
                )
        except Exception:
            pass

    def _on_cmb_bar_selection_changed(self, sender, args):
        self._sync_bar_type_tooltip()
        self._sync_lap_tb_from_selected_bar_diam()
        self._refresh_max_length_ui_from_current_selection()

    def _get_selected_bar_nominal_diameter_mm(self):
        cmb = self._win.FindName("CmbBarType")
        if cmb is None or not self._entries:
            return 12.0
        try:
            idx = int(cmb.SelectedIndex)
        except Exception:
            idx = 0
        if idx < 0 or idx >= len(self._entries):
            idx = 0
        b_pick, _ = self._entries[idx]
        if b_pick is None:
            return 12.0
        d_mm = _rebar_nominal_diameter_mm(b_pick)
        return float(d_mm) if d_mm else 12.0

    def _sync_lap_tb_from_selected_bar_diam(self):
        cmb = self._win.FindName("CmbBarType")
        tlap = self._win.FindName("TxtLapMm")
        if cmb is None or tlap is None or not self._entries:
            return
        try:
            idx = int(cmb.SelectedIndex)
        except Exception:
            idx = 0
        if idx < 0 or idx >= len(self._entries):
            idx = 0
        b_pick, _ = self._entries[idx]
        if b_pick is None:
            d_mm = 12.0
        else:
            d_mm = _rebar_nominal_diameter_mm(b_pick)
            if d_mm is None:
                d_mm = 12.0
        try:
            from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm

            v = traslape_mm_from_nominal_diameter_mm(float(d_mm), None)
            if v is not None:
                tlap.Text = unicode(int(round(float(v))))
        except Exception:
            pass

    def _set_max_bar_length_controls_visible(self, visible):
        pnl = self._win.FindName("PnlMaxBarLength")
        if pnl is None:
            return
        try:
            from System.Windows import Visibility
            pnl.Visibility = Visibility.Visible if bool(visible) else Visibility.Collapsed
        except Exception:
            try:
                pnl.Visibility = "Visible" if bool(visible) else "Collapsed"
            except Exception:
                pass

    def _is_max_bar_length_controls_visible(self):
        pnl = self._win.FindName("PnlMaxBarLength")
        if pnl is None:
            return False
        try:
            from System.Windows import Visibility
            return pnl.Visibility == Visibility.Visible
        except Exception:
            try:
                return unicode(pnl.Visibility) == u"Visible"
            except Exception:
                return False

    def _focus_max_bar_length_input(self):
        txt = self._win.FindName("TxtMaxBarLength")
        if txt is None:
            return
        try:
            txt.Focus()
            txt.SelectAll()
        except Exception:
            pass

    def _get_selected_forced_bar_type_id(self):
        cmb = self._win.FindName("CmbBarType")
        try:
            idx = int(cmb.SelectedIndex) if cmb is not None else 0
        except Exception:
            idx = 0
        if idx < 0 or idx >= len(self._entries):
            idx = 0
        if idx < 0 or idx >= len(self._entries):
            return None
        b_pick, _lbl = self._entries[idx]
        if b_pick is None:
            return None
        try:
            return b_pick.Id
        except Exception:
            return None

    def _refresh_max_length_ui_from_current_selection(self):
        if self._document is None or self._host is None or not self._refs:
            self._set_max_bar_length_controls_visible(False)
            self._update_troceo_aviso_ui(False)
            self._max_length_confirmed = False
            return
        try:
            from enfierrado_shaft_hashtag import evaluar_si_excede_12m_en_shaft

            bid = self._get_selected_forced_bar_type_id()
            excede_12m, max_detect_mm, err_eval = evaluar_si_excede_12m_en_shaft(
                self._document,
                self._host,
                self._refs,
                forced_bar_type_id=bid,
                ignore_empotramientos=False,
                empotramiento_adaptivo_extremos=True,
            )
            if err_eval:
                self._set_estado(u"No se pudo evaluar largo previo: {0}".format(err_eval))
                self._update_troceo_aviso_ui(False)
                return
            self._set_max_bar_length_controls_visible(bool(excede_12m))
            if excede_12m:
                self._sync_lap_tb_from_selected_bar_diam()
                self._focus_max_bar_length_input()
                if max_detect_mm is not None:
                    self._update_troceo_aviso_ui(
                        True,
                        u"La barra ({0:.0f} mm) excede el máximo (12 000 mm). Defina tramos.".format(
                            float(max_detect_mm)
                        ),
                    )
                else:
                    self._update_troceo_aviso_ui(
                        True,
                        u"La barra excede el máximo (12 000 mm). Defina tramos.",
                    )
                self._set_estado(u"")
            else:
                self._update_troceo_aviso_ui(False)
                self._max_length_confirmed = False
        except Exception as ex:
            self._set_estado(u"No se pudo evaluar largo máximo: {0}".format(ex))

    def _read_max_bar_length_from_form(self):
        txt = self._win.FindName("TxtMaxBarLength")
        if txt is None:
            return None, u"No se encontró el campo de largo máximo por tramo."
        raw = unicode(txt.Text or u"").strip()
        if not raw:
            return None, u"Ingrese el largo máximo por tramo (mm)."
        if not raw.isdigit():
            return None, u"El largo máximo debe ser un número entero en milímetros."
        mm = float(raw)
        if mm < float(_MAX_BAR_LENGTH_MIN_MM) or mm > float(_MAX_BAR_LENGTH_RECOMMENDED_MAX_MM):
            return None, u"Ingrese un valor entre {0:.0f} y {1:.0f} mm.".format(
                float(_MAX_BAR_LENGTH_MIN_MM),
                float(_MAX_BAR_LENGTH_RECOMMENDED_MAX_MM),
            )
        return mm, None

    def _sync_bar_type_tooltip(self):
        cmb = self._win.FindName("CmbBarType")
        if cmb is None:
            return
        try:
            idx = int(cmb.SelectedIndex)
            if idx < 0 or idx >= len(self._entries):
                cmb.ToolTip = None
                return
            b_pick, lbl = self._entries[idx]
            if b_pick is None:
                cmb.ToolTip = u"Automático: tipo más cercano a 12 mm nominal."
                return
            try:
                nm = unicode(b_pick.Name or u"").strip()
            except Exception:
                nm = u""
            if nm and nm.lower() not in lbl.lower():
                cmb.ToolTip = u"{0} — {1}".format(lbl, nm)
            else:
                cmb.ToolTip = lbl
        except Exception:
            pass

    def _refresh_capas_suffix_label(self):
        tb = self._win.FindName("TxtNumCapasBordeLosa")
        lbl = self._win.FindName("TbCapasSuffixBordeLosa")
        if lbl is None:
            return
        n = _CAPAS_OPCIONES_MIN
        if tb is not None:
            try:
                s = unicode(tb.Text).strip()
                if s:
                    n = int(float(s.replace(u",", u".")))
            except Exception:
                n = _CAPAS_OPCIONES_MIN
        try:
            lbl.Text = u"Capa" if n == 1 else u"Capas"
        except Exception:
            pass

    def _wire_capas_spinner_once(self):
        if getattr(self, "_capas_spinner_hooked", False):
            return
        tb = self._win.FindName("TxtNumCapasBordeLosa")
        bu = self._win.FindName("BtnNumCapasBordeLosaUp")
        bd = self._win.FindName("BtnNumCapasBordeLosaDown")
        if tb is None:
            return
        self._capas_spinner_hooked = True
        from System.Windows import RoutedEventHandler, DataObject, DataObjectPastingEventHandler
        from System.Windows.Input import TextCompositionEventHandler

        def _is_digits_only(s):
            if s is None:
                return False
            t = unicode(s)
            if not t:
                return False
            for ch in t:
                if ch < u"0" or ch > u"9":
                    return False
            return True

        def _preview_text_input(sender, e):
            try:
                e.Handled = not _is_digits_only(e.Text)
            except Exception:
                pass

        def _on_pasting(sender, e):
            try:
                txt = None
                if e.SourceDataObject is not None and e.SourceDataObject.GetDataPresent("Text"):
                    txt = e.SourceDataObject.GetData("Text")
                if not _is_digits_only(txt):
                    e.CancelCommand()
            except Exception:
                try:
                    e.CancelCommand()
                except Exception:
                    pass

        def _lf_capas(s, e):
            _normalize_capas_borde_losa_textbox(tb)
            try:
                self._refresh_capas_suffix_label()
            except Exception:
                pass
            try:
                self._refresh_max_length_ui_from_current_selection()
            except Exception:
                pass

        def _tc_capas(s, e):
            try:
                self._refresh_capas_suffix_label()
            except Exception:
                pass
            try:
                self._refresh_max_length_ui_from_current_selection()
            except Exception:
                pass

        try:
            tb.LostFocus += RoutedEventHandler(_lf_capas)
            tb.TextChanged += RoutedEventHandler(_tc_capas)
            tb.PreviewTextInput += TextCompositionEventHandler(_preview_text_input)
            DataObject.AddPastingHandler(tb, DataObjectPastingEventHandler(_on_pasting))
        except Exception:
            pass

        if bu is not None:
            def _cap_up(s, e):
                _bump_capas_borde_losa_textbox(tb, 1)
                try:
                    self._refresh_capas_suffix_label()
                except Exception:
                    pass
                try:
                    self._refresh_max_length_ui_from_current_selection()
                except Exception:
                    pass

            try:
                bu.Click += RoutedEventHandler(_cap_up)
            except Exception:
                pass
        if bd is not None:
            def _cap_down(s, e):
                _bump_capas_borde_losa_textbox(tb, -1)
                try:
                    self._refresh_capas_suffix_label()
                except Exception:
                    pass
                try:
                    self._refresh_max_length_ui_from_current_selection()
                except Exception:
                    pass

            try:
                bd.Click += RoutedEventHandler(_cap_down)
            except Exception:
                pass

        try:
            self._refresh_capas_suffix_label()
        except Exception:
            pass

    def _reset_capas_spinner(self):
        tb = self._win.FindName("TxtNumCapasBordeLosa")
        if tb is None:
            return
        try:
            tb.Text = unicode(_CAPAS_OPCIONES_MIN)
        except Exception:
            pass
        try:
            self._refresh_capas_suffix_label()
        except Exception:
            pass

    def _read_n_capas_from_form(self):
        tb = self._win.FindName("TxtNumCapasBordeLosa")
        if tb is None:
            return _CAPAS_OPCIONES_MIN
        try:
            _normalize_capas_borde_losa_textbox(tb)
            try:
                self._refresh_capas_suffix_label()
            except Exception:
                pass
            s = unicode(tb.Text).strip()
            if not s:
                return _CAPAS_OPCIONES_MIN
            return int(s)
        except Exception:
            return _CAPAS_OPCIONES_MIN

    def _wire_barras_set_spinner_once(self):
        if getattr(self, "_barras_set_spinner_hooked", False):
            return
        tb = self._win.FindName("TxtCantidadBarrasSet")
        bu = self._win.FindName("BtnCantidadBarrasSetUp")
        bd = self._win.FindName("BtnCantidadBarrasSetDown")
        if tb is None:
            return
        self._barras_set_spinner_hooked = True
        from System.Windows import RoutedEventHandler, DataObject, DataObjectPastingEventHandler
        from System.Windows.Input import TextCompositionEventHandler

        def _is_digits_only(s):
            if s is None:
                return False
            t = unicode(s)
            if not t:
                return False
            for ch in t:
                if ch < u"0" or ch > u"9":
                    return False
            return True

        def _preview_text_input(sender, e):
            try:
                e.Handled = not _is_digits_only(e.Text)
            except Exception:
                pass

        def _on_pasting(sender, e):
            try:
                txt = None
                if e.SourceDataObject is not None and e.SourceDataObject.GetDataPresent("Text"):
                    txt = e.SourceDataObject.GetData("Text")
                if not _is_digits_only(txt):
                    e.CancelCommand()
            except Exception:
                try:
                    e.CancelCommand()
                except Exception:
                    pass

        def _lf_bs(s, e):
            _normalize_barras_set_borde_losa_textbox(tb)
            try:
                self._refresh_max_length_ui_from_current_selection()
            except Exception:
                pass

        def _tc_bs(s, e):
            try:
                self._refresh_max_length_ui_from_current_selection()
            except Exception:
                pass

        try:
            tb.LostFocus += RoutedEventHandler(_lf_bs)
            tb.TextChanged += RoutedEventHandler(_tc_bs)
            tb.PreviewTextInput += TextCompositionEventHandler(_preview_text_input)
            DataObject.AddPastingHandler(tb, DataObjectPastingEventHandler(_on_pasting))
        except Exception:
            pass

        if bu is not None:
            def _bs_up(s, e):
                _bump_barras_set_borde_losa_textbox(tb, 1)
                try:
                    self._refresh_max_length_ui_from_current_selection()
                except Exception:
                    pass

            try:
                bu.Click += RoutedEventHandler(_bs_up)
            except Exception:
                pass
        if bd is not None:
            def _bs_down(s, e):
                _bump_barras_set_borde_losa_textbox(tb, -1)
                try:
                    self._refresh_max_length_ui_from_current_selection()
                except Exception:
                    pass

            try:
                bd.Click += RoutedEventHandler(_bs_down)
            except Exception:
                pass

    def _reset_barras_set_spinner(self):
        tb = self._win.FindName("TxtCantidadBarrasSet")
        if tb is None:
            return
        try:
            tb.Text = unicode(_DEFAULT_BARRAS_SET_UI)
        except Exception:
            pass

    def _read_n_barras_set_from_form(self):
        tb = self._win.FindName("TxtCantidadBarrasSet")
        if tb is None:
            return _DEFAULT_BARRAS_SET_UI
        try:
            _normalize_barras_set_borde_losa_textbox(tb)
            s = unicode(tb.Text).strip()
            if not s:
                return _DEFAULT_BARRAS_SET_UI
            return int(s)
        except Exception:
            return _DEFAULT_BARRAS_SET_UI

    def _wire_commands(self, RoutedEventHandler, ApplicationCommands, CommandBinding, KeyBinding, Key, ModifierKeys):
        def _on_close_cmd(sender, e):
            try:
                self._close_with_fade()
            except Exception:
                pass

        try:
            self._win.CommandBindings.Add(
                CommandBinding(ApplicationCommands.Close, RoutedEventHandler(_on_close_cmd))
            )
            self._win.InputBindings.Add(
                KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None)
            )
        except Exception:
            pass

    def _close_with_fade(self):
        """Cierra la ventana con fade-out y fallback seguro."""
        if getattr(self, "_is_closing_with_fade", False):
            return
        self._is_closing_with_fade = True
        try:
            from System import TimeSpan
            from System.Windows import Duration
            from System.Windows.Media.Animation import DoubleAnimation, QuadraticEase, EasingMode

            try:
                self._win.BeginAnimation(self._win.OpacityProperty, None)
                self._win.BeginAnimation(self._win.TopProperty, None)
            except Exception:
                pass

            ease_out = QuadraticEase()
            ease_out.EasingMode = EasingMode.EaseInOut

            opacity_anim = DoubleAnimation()
            opacity_anim.From = float(self._win.Opacity)
            opacity_anim.To = 0.0
            opacity_anim.Duration = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_CLOSE_MS)))
            opacity_anim.EasingFunction = ease_out

            current_top = float(self._win.Top)
            target_top = current_top + float(_WINDOW_SLIDE_PX)
            top_anim = DoubleAnimation()
            top_anim.From = current_top
            top_anim.To = target_top
            top_anim.Duration = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_CLOSE_MS)))
            top_anim.EasingFunction = ease_out

            def _on_done(sender, args):
                try:
                    self._win.Close()
                except Exception:
                    pass

            opacity_anim.Completed += _on_done
            self._win.BeginAnimation(self._win.OpacityProperty, opacity_anim)
            self._win.BeginAnimation(self._win.TopProperty, top_anim)
        except Exception:
            self._is_closing_with_fade = False
            try:
                self._win.Close()
            except Exception:
                pass

    def _show_with_fade(self):
        """Muestra la ventana con fade-in y fallback seguro."""
        try:
            from System import TimeSpan
            from System.Windows import Duration
            from System.Windows.Media.Animation import DoubleAnimation, QuadraticEase, EasingMode

            # Evita acumulación de animaciones al abrir/cerrar repetidamente.
            try:
                self._win.BeginAnimation(self._win.OpacityProperty, None)
                self._win.BeginAnimation(self._win.TopProperty, None)
            except Exception:
                pass

            self._win.Opacity = 0.0
            if not self._win.IsVisible:
                self._win.Show()
            # Con posición Manual (vista activa), Left/Top ya están fijados antes del Show.
            # Forzamos layout antes de calcular Top para que el slide-in no se omita.
            try:
                self._win.UpdateLayout()
            except Exception:
                pass

            try:
                self._base_top = float(self._win.Top)
            except Exception:
                if self._base_top is None:
                    self._base_top = 0.0

            start_top = float(self._base_top) + float(_WINDOW_SLIDE_PX)
            try:
                self._win.Top = start_top
            except Exception:
                pass

            ease_in = QuadraticEase()
            ease_in.EasingMode = EasingMode.EaseInOut

            opacity_anim = DoubleAnimation()
            opacity_anim.From = 0.0
            opacity_anim.To = 1.0
            opacity_anim.Duration = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_OPEN_MS)))
            opacity_anim.EasingFunction = ease_in

            top_anim = DoubleAnimation()
            top_anim.From = start_top
            top_anim.To = float(self._base_top)
            top_anim.Duration = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_OPEN_MS)))
            top_anim.EasingFunction = ease_in

            self._win.BeginAnimation(self._win.OpacityProperty, opacity_anim)
            self._win.BeginAnimation(self._win.TopProperty, top_anim)
            self._is_closing_with_fade = False
            self._win.Activate()
        except Exception:
            try:
                self._win.Opacity = 1.0
            except Exception:
                pass
            try:
                if not self._win.IsVisible:
                    self._win.Show()
                self._is_closing_with_fade = False
                self._win.Activate()
            except Exception:
                pass

    def _set_seleccion_line(self, msg):
        try:
            txt = self._win.FindName("TxtSeleccionLine")
            if txt is not None:
                txt.Text = msg or u""
        except Exception:
            pass

    def _update_troceo_aviso_ui(self, visible, text=None):
        br = self._win.FindName("BorderTroceoAviso")
        tba = self._win.FindName("TxtTroceoAviso")
        if br is None:
            return
        try:
            from System.Windows import Visibility

            br.Visibility = Visibility.Visible if visible else Visibility.Collapsed
        except Exception:
            try:
                br.Visibility = "Visible" if visible else "Collapsed"
            except Exception:
                pass
        if tba is None:
            return
        if not visible or text is None:
            try:
                tba.Text = u""
            except Exception:
                pass
            return
        try:
            from System.Windows.Documents import Run
            from System.Windows import FontWeights

            tba.Inlines.Clear()
            r0 = Run(u"Alerta: ")
            r0.FontWeight = FontWeights.Bold
            tba.Inlines.Add(r0)
            tba.Inlines.Add(Run(text))
        except Exception:
            try:
                tba.Text = u"Alerta: {0}".format(text)
            except Exception:
                pass

    def _set_estado(self, msg):
        try:
            txt = self._win.FindName("TxtEstado")
            if txt is not None:
                s = msg or u""
                txt.Text = s
                try:
                    from System.Windows import Visibility

                    txt.Visibility = (
                        Visibility.Visible if bool(s.strip()) else Visibility.Collapsed
                    )
                except Exception:
                    pass
        except Exception:
            pass

    def _actualizar_info_host_caras(self):
        pass

    def _limpiar_seleccion_host_caras(self):
        self._host = None
        self._refs = None
        try:
            self._max_length_confirmed = False
        except Exception:
            pass
        self._set_max_bar_length_controls_visible(False)
        self._set_seleccion_line(u"")
        self._actualizar_info_host_caras()
        self._refresh_max_length_ui_from_current_selection()
        try:
            from seleccion_caras_elemento import _asignar_resultado_modulo

            _asignar_resultado_modulo(None, None)
        except Exception:
            pass

    def _cargar_combos(self):
        doc = self._document
        if doc is None:
            return
        entries, err = _build_bar_type_entries(doc)
        cmb = self._win.FindName("CmbBarType")
        if err:
            self._entries = []
            if cmb is not None:
                cmb.Items.Clear()
            self._set_estado(err)
            return
        self._entries = list(entries)
        if cmb is None:
            return
        cmb.Items.Clear()
        cmb.IsEditable = False
        for _bt, lbl in self._entries:
            cmb.Items.Add(lbl)
        sel_idx = 0
        for i, (b, _) in enumerate(self._entries):
            if b is None:
                sel_idx = i
                break
        cmb.SelectedIndex = sel_idx
        try:
            if not getattr(self, "_bar_sel_hooked", False):
                # SelectionChanged exige SelectionChangedEventHandler; RoutedEventHandler
                # no invoca el callback de forma fiable en ComboBox (IronPython / WPF).
                from System.Windows.Controls import SelectionChangedEventHandler

                cmb.SelectionChanged += SelectionChangedEventHandler(
                    self._on_cmb_bar_selection_changed
                )
                self._bar_sel_hooked = True
        except Exception:
            pass
        self._sync_bar_type_tooltip()
        self._sync_lap_tb_from_selected_bar_diam()

    def _on_seleccionar(self, sender, args):
        try:
            self._win.Hide()
        except Exception:
            pass
        self._seleccion_event.Raise()

    def _on_generar(self, sender, args):
        if self._host is None or not self._refs:
            _task_dialog_show(
                _TOOL_TASK_DIALOG_TITLE,
                u"Seleccione primero el host y las caras.",
                self._win,
            )
            return
        n_capas = self._read_n_capas_from_form()
        n_capas = max(
            _CAPAS_OPCIONES_MIN,
            min(_CAPAS_OPCIONES_MAX, int(n_capas)),
        )

        n_barras_set = self._read_n_barras_set_from_form()
        n_barras_set = max(
            _BARRAS_SET_MIN,
            min(_BARRAS_SET_MAX, int(n_barras_set)),
        )
        bid = self._get_selected_forced_bar_type_id()
        self._generar_handler.lap_length_mm = None
        place_lap_details = True
        lap_detail_symbol_id, lap_detail_err = _find_fixed_lap_detail_symbol_id(self._document)
        if lap_detail_symbol_id is None:
            # No bloquear la creación de barras si falta el símbolo fijo.
            place_lap_details = False
            self._set_estado(
                (lap_detail_err or u"No se encontró el Detail Component fijo para empalmes.")
                + u" Se continuará sin detalle de empalmes."
            )
        elif lap_detail_err:
            self._set_estado(lap_detail_err)

        max_bar_length_mm = float(self._default_max_bar_length_mm)
        excede_12m = False
        max_detect_mm = None
        try:
            from enfierrado_shaft_hashtag import evaluar_si_excede_12m_en_shaft

            excede_12m, max_detect_mm, err_eval = evaluar_si_excede_12m_en_shaft(
                self._document,
                self._host,
                self._refs,
                forced_bar_type_id=bid,
                ignore_empotramientos=False,
                empotramiento_adaptivo_extremos=True,
            )
            if err_eval:
                self._set_estado(u"No se pudo evaluar largo previo: {0}".format(err_eval))
                return
        except Exception as ex:
            self._set_estado(u"No se pudo evaluar largo máximo: {0}".format(ex))
            return

        if excede_12m:
            was_visible = self._is_max_bar_length_controls_visible()
            self._set_max_bar_length_controls_visible(True)
            if not was_visible:
                self._focus_max_bar_length_input()
                if max_detect_mm is not None:
                    self._update_troceo_aviso_ui(
                        True,
                        u"La barra ({0:.0f} mm) excede el máximo (12 000 mm). "
                        u"Ingrese el largo máximo por tramo y vuelva a presionar «Colocar armaduras».".format(
                            float(max_detect_mm)
                        ),
                    )
                else:
                    self._update_troceo_aviso_ui(
                        True,
                        u"La barra excede el máximo (12 000 mm). "
                        u"Ingrese el largo máximo por tramo y vuelva a presionar «Colocar armaduras».",
                    )
                self._set_estado(u"")
                return
            max_bar_length_mm, err_max = self._read_max_bar_length_from_form()
            if err_max:
                if max_detect_mm is not None:
                    self._set_estado(
                        u"La barra detectada ({0:.0f} mm) excede 12000 mm. {1}".format(
                            float(max_detect_mm),
                            err_max,
                        )
                    )
                else:
                    self._set_estado(
                        u"La barra excede 12000 mm. {0}".format(err_max)
                    )
                return
            lap_tb = self._win.FindName("TxtLapMm")
            d_bar_mm = self._get_selected_bar_nominal_diameter_mm()
            lap_mm = _borde_losa_traslape_mm(d_bar_mm, lap_tb)
            if max_bar_length_mm <= lap_mm + 1.0:
                _task_dialog_show(
                    _TOOL_TASK_DIALOG_TITLE,
                    u"El largo máximo por tramo debe ser mayor que el empalme.",
                    self._win,
                )
                return
            self._generar_handler.lap_length_mm = lap_mm
            self._default_max_bar_length_mm = float(max_bar_length_mm)
            self._max_length_confirmed = True
        else:
            self._set_max_bar_length_controls_visible(False)
            self._update_troceo_aviso_ui(False)
            self._max_length_confirmed = False

        self._generar_handler.n_capas = n_capas
        self._generar_handler.n_barras_set = n_barras_set
        self._generar_handler.forced_bar_type_id = bid
        self._generar_handler.max_bar_length_mm = max_bar_length_mm
        self._generar_handler.place_lap_details = bool(place_lap_details)
        self._generar_handler.lap_detail_symbol_id = lap_detail_symbol_id
        self._generar_event.Raise()

    def show(self):
        uidoc = self._revit.ActiveUIDocument
        if uidoc is None:
            _task_dialog_show(
                _TOOL_TASK_DIALOG_TITLE,
                u"No hay documento activo.",
                self._win,
            )
            return
        self._document = uidoc.Document
        hwnd = None
        try:
            from System.Windows.Interop import WindowInteropHelper

            hwnd = _revit_main_hwnd(self._revit.Application)
            if hwnd:
                helper = WindowInteropHelper(self._win)
                helper.Owner = hwnd
        except Exception:
            pass
        position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        self._reset_capas_spinner()
        self._reset_barras_set_spinner()
        self._cargar_combos()
        self._actualizar_info_host_caras()
        self._set_max_bar_length_controls_visible(False)
        self._set_seleccion_line(u"")
        self._update_troceo_aviso_ui(False)
        self._set_estado(u"")
        self._show_with_fade()


def run_pyrevit(revit):
    _scripts_dir = os.path.dirname(os.path.abspath(__file__))
    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    existing = _get_active_window()
    if existing is not None:
        try:
            from System.Windows import WindowState

            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
        except Exception:
            pass
        try:
            existing.Show()
        except Exception:
            pass
        try:
            existing.Activate()
            existing.Focus()
        except Exception:
            pass
        _task_dialog_show(
            _TOOL_TASK_DIALOG_TITLE,
            u"La herramienta ya está en ejecución.",
            existing,
        )
        return

    w = EnfierradoShaftPasadaWindow(revit)
    _set_active_window(w._win)
    w.show()


def _get_active_window():
    try:
        win = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        # Si la referencia quedó stale, limpiar y devolver None.
        _ = win.Title
    except Exception:
        _clear_active_window()
        return None
    try:
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            _clear_active_window()
            return None
    except Exception:
        pass
    return win


def _set_active_window(win):
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, win)
    except Exception:
        pass


def _clear_active_window():
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass
