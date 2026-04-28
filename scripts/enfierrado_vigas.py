# -*- coding: utf-8 -*-
"""
Armadura vigas — formulario (misma línea visual que enfierrado fundación aislada).

- Estilos WPF y ventana sin borde.
- Selección múltiple: Structural Framing, Structural Columns y Walls.
- «Colocar armadura»: cara sup. (-25 mm n, eje V = ancho/2−25 mm), unificación, +/-2 m/ext.,
  troceo si aplica, −25 mm por extremo en eje; ``ModelCurve`` opcional (checkbox *Model lines*).
- **Armadura superior**: **Nº de capas** (1–3, ▲▼); una sola fila de encabezados **Cantidad** / **ø Diámetro (mm)**
  compartida por **1ªC.–3ªC.**; diámetros en ``ComboBox`` con valores estándar (8–32 mm), resueltos a ``RebarBarType``
  en el proyecto. Filas **2ªC./3ªC.** se muestran u ocultan según **Nº de capas** (``Collapsed`` si no aplican).
  **Suple superior**: **Nº de capas** 1–2 (tras el título); la 1.ª fila suple es siempre
  ``(N_main+1)ªC.`` y la 2.ª ``(N_main+2)ªC.`` si aplica.
  empalmes si L>12 m.
- Armadura **inferior**: misma disposición en tabla (1ªC.–3ªC., suple inferior);
  eje proyectado a la cara inferior (mismo troceo/recortes).
- **Barras laterales** (columna derecha): cantidad y diámetro en UI; al colocar **cara superior**
  y/o **inferior** activa se crean ``Rebar`` por tramo en esa cara (curva guía tras última capa
  + offset, ``norm = −n`` con respaldos, ``SetLayoutAsFixedNumber``, canto − 2×offset respecto la cara).
- **Estribos** (tras armadura inferior en la columna principal): filas Ext./Cent. con separación (mm) y diámetro;
  un solo **tipo** (Simple / Doble / Triple) bajo la tabla; opción activa por defecto; resumen al colocar;
  generación pendiente de enlace.
- Bloque **Empalmes**: bajo **Seleccionar en Modelo**, visible si el trazo superior **o inferior** estimado
  supera 12 m; **troceo automático** al colocar si L &gt; 12 m (sin checkbox). Vigas Structural Framing
  para planos de empalme (sin listado de Id en el formulario). Si el trazo &gt; 12 m, **Colocar armadura**
  exige al menos una viga de empalme definida; si no, ``TaskDialog`` y no se ejecuta la colocación.
"""

import math
import os
import re
import sys
import weakref
import clr
import System

_scripts_dir = os.path.dirname(os.path.abspath(__file__))
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ISelectionFilter

from barras_bordes_losa_gancho_empotramiento import (
    _rebar_nominal_diameter_mm,
    element_id_to_int,
    _task_dialog_show,
)
from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_paths import get_logo_paths

_APPDOMAIN_WINDOW_KEY = "BIMTools.EnfierradoVigas.ActiveWindow"

_ALLOWED_SELECTION_CAT_IDS = frozenset(
    (
        int(BuiltInCategory.OST_StructuralFraming),
        int(BuiltInCategory.OST_StructuralColumns),
        int(BuiltInCategory.OST_Walls),
    )
)

# Valor inicial y rango del campo numérico «cantidad de barras» (por capa / suple).
_CANTIDAD_BARRAS_MIN = 1
_CANTIDAD_BARRAS_MAX = 99
_CANTIDAD_BARRAS_DEFAULT_TXT = u"2"
_CANTIDAD_SPINNER_TRIPLES = (
    (u"CmbSupCant", u"BtnCmbSupCantUp", u"BtnCmbSupCantDown"),
    (u"CmbSup2Cant", u"BtnCmbSup2CantUp", u"BtnCmbSup2CantDown"),
    (u"CmbSup3Cant", u"BtnCmbSup3CantUp", u"BtnCmbSup3CantDown"),
    (u"CmbSupleCant", u"BtnCmbSupleCantUp", u"BtnCmbSupleCantDown"),
    (u"CmbSuple2Cant", u"BtnCmbSuple2CantUp", u"BtnCmbSuple2CantDown"),
    (u"CmbInfCant", u"BtnCmbInfCantUp", u"BtnCmbInfCantDown"),
    (u"CmbInf2Cant", u"BtnCmbInf2CantUp", u"BtnCmbInf2CantDown"),
    (u"CmbInf3Cant", u"BtnCmbInf3CantUp", u"BtnCmbInf3CantDown"),
    (u"CmbSupleInfCant", u"BtnCmbSupleInfCantUp", u"BtnCmbSupleInfCantDown"),
    (u"CmbLatCant", u"BtnCmbLatCantUp", u"BtnCmbLatCantDown"),
)

# Separación entre estribos (mm): campos Extremos / Centrales.
_ESTRIBO_SEPARACION_MM_MIN = 50
_ESTRIBO_SEPARACION_MM_MAX = 999
_ESTRIBO_SEPARACION_DEFAULT_TXT = u"200"
_ESTRIBO_SEP_SPINNER_TRIPLES = (
    (u"TxtEstriboExtSep", u"BtnEstriboExtSepUp", u"BtnEstriboExtSepDown"),
    (u"TxtEstriboCentSep", u"BtnEstriboCentSepUp", u"BtnEstriboCentSepDown"),
)

_CAPAS_ARMADURA_MIN = 1
_CAPAS_ARMADURA_MAX = 3
_CAPAS_DEFAULT_TXT = u"1"
_CAPAS_SUPLE_MIN = 1
_CAPAS_SUPLE_MAX = 2
_CAPAS_SUPLE_DEFAULT_TXT = u"1"

# Diámetros estándar (mm) en combos; resolución a ``RebarBarType`` vía documento.
_DIAMETROS_ESTANDAR_MM = (8, 10, 12, 16, 18, 22, 25, 28, 32)

def _etiquetas_diametro_estandar():
    return [u"ø{0} mm".format(mm) for mm in _DIAMETROS_ESTANDAR_MM]


def _param_double_mm(elem, bip):
    """Convierte un parámetro de longitud de ``elem`` a mm, o ``None``."""
    if elem is None:
        return None
    try:
        from Autodesk.Revit.DB import StorageType, UnitTypeId, UnitUtils

        p = elem.get_Parameter(bip)
        if (
            p is None
            or not p.HasValue
            or p.StorageType != StorageType.Double
        ):
            return None
        return float(
            UnitUtils.ConvertFromInternalUnits(
                p.AsDouble(), UnitTypeId.Millimeters
            )
        )
    except Exception:
        return None


def _altura_viga_structural_mm(elem, document=None):
    """
    Canto (altura de sección) de una viga en mm para reglas de armado.

    Prioriza ``armadura_vigas_capas._read_width_depth_ft`` (misma convención que
    el resto de herramientas: parámetros *Height* / *Depth* del tipo y bbox con las
    dos dimensiones transversales, no la luz). ``STRUCTURAL_DEPTH`` en instancia
    suele alinearse a un lado de la sección y puede dar ~500 mm cuando el canto
    real es 800 mm.
    """
    if elem is None:
        return None
    if document is not None:
        try:
            from armadura_vigas_capas import _read_width_depth_ft
            from Autodesk.Revit.DB import LocationCurve

            loc = getattr(elem, "Location", None)
            if isinstance(loc, LocationCurve):
                crv = loc.Curve
                if crv is not None:
                    _w_ft, d_ft = _read_width_depth_ft(document, elem, crv)
                    if d_ft is not None and float(d_ft) > 0.0:
                        d_mm = float(d_ft) * 304.8
                        if d_mm > 0.5:
                            return d_mm
        except Exception:
            pass
    try:
        _bip_depth = BuiltInParameter.STRUCTURAL_DEPTH
    except AttributeError:
        _bip_depth = None
    if _bip_depth is not None:
        v = _param_double_mm(elem, _bip_depth)
        if v is not None and v > 0.5:
            return v
    if document is not None:
        try:
            et = document.GetElement(elem.GetTypeId())
        except Exception:
            et = None
        if et is not None:
            for nm in (
                u"Height",
                u"Depth",
                u"Altura",
                u"Profundidad",
                u"Canto",
                u"h",
                u"H",
                u"d",
            ):
                try:
                    p = et.LookupParameter(nm)
                    if p is None or not p.HasValue:
                        continue
                    from Autodesk.Revit.DB import StorageType, UnitTypeId, UnitUtils

                    if p.StorageType != StorageType.Double:
                        continue
                    return float(
                        UnitUtils.ConvertFromInternalUnits(
                            p.AsDouble(), UnitTypeId.Millimeters
                        )
                    )
                except Exception:
                    continue
    for nm in (u"Depth", u"Structural Depth", u"Profundidad", u"Altura", u"Canto"):
        try:
            p = elem.LookupParameter(nm)
            if p is None or not p.HasValue:
                continue
            from Autodesk.Revit.DB import StorageType, UnitTypeId, UnitUtils

            if p.StorageType != StorageType.Double:
                continue
            return float(
                UnitUtils.ConvertFromInternalUnits(
                    p.AsDouble(), UnitTypeId.Millimeters
                )
            )
        except Exception:
            continue
    try:
        bb = elem.get_BoundingBox(None)
        if bb is None:
            return None
        dv = bb.Max - bb.Min
        from Autodesk.Revit.DB import UnitTypeId, UnitUtils

        dx = abs(float(dv.X))
        dy = abs(float(dv.Y))
        dz = abs(float(dv.Z))
        dims = sorted([dx, dy, dz], reverse=True)
        if len(dims) >= 3:
            sec_ft = max(dims[1], dims[2])
        else:
            sec_ft = dims[-1]
        return float(
            UnitUtils.ConvertFromInternalUnits(sec_ft, UnitTypeId.Millimeters)
        )
    except Exception:
        return None


def _cantidad_laterales_inicial_desde_altura_mm(h_mm):
    """
    Cantidad sugerida: ``ceil(h_mm/200) - 1`` acotada a la UI (mín. 1).
    Ej.: 720 mm → 3,6 → entero superior 4 → menos 1 → **3**.
    """
    if h_mm is None or h_mm <= 0:
        return max(_CANTIDAD_BARRAS_MIN, int(_CANTIDAD_BARRAS_DEFAULT_TXT))
    cociente = float(h_mm) / 200.0
    n = int(math.ceil(cociente)) - 1
    return max(_CANTIDAD_BARRAS_MIN, min(_CANTIDAD_BARRAS_MAX, n))


def _normalize_capas_textbox(tb):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = _CAPAS_DEFAULT_TXT
            return
        n = int(float(s.replace(u",", u".")))
        n = max(_CAPAS_ARMADURA_MIN, min(_CAPAS_ARMADURA_MAX, n))
        tb.Text = unicode(n)
    except Exception:
        tb.Text = _CAPAS_DEFAULT_TXT


def _bump_capas_textbox(tb, delta):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if s:
            n = int(float(s.replace(u",", u".")))
        else:
            n = int(_CAPAS_DEFAULT_TXT)
    except Exception:
        n = int(_CAPAS_DEFAULT_TXT)
    n = max(
        _CAPAS_ARMADURA_MIN,
        min(_CAPAS_ARMADURA_MAX, n + int(delta)),
    )
    tb.Text = unicode(n)


def _normalize_capas_suple_textbox(tb):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = _CAPAS_SUPLE_DEFAULT_TXT
            return
        n = int(float(s.replace(u",", u".")))
        n = max(
            _CAPAS_SUPLE_MIN,
            min(_CAPAS_SUPLE_MAX, n),
        )
        tb.Text = unicode(n)
    except Exception:
        tb.Text = _CAPAS_SUPLE_DEFAULT_TXT


def _bump_capas_suple_textbox(tb, delta):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if s:
            n = int(float(s.replace(u",", u".")))
        else:
            n = int(_CAPAS_SUPLE_DEFAULT_TXT)
    except Exception:
        n = int(_CAPAS_SUPLE_DEFAULT_TXT)
    n = max(
        _CAPAS_SUPLE_MIN,
        min(_CAPAS_SUPLE_MAX, n + int(delta)),
    )
    tb.Text = unicode(n)


def _normalize_cantidad_textbox(tb):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = _CANTIDAD_BARRAS_DEFAULT_TXT
            return
        n = int(float(s.replace(u",", u".")))
        n = max(_CANTIDAD_BARRAS_MIN, min(_CANTIDAD_BARRAS_MAX, n))
        tb.Text = unicode(n)
    except Exception:
        tb.Text = _CANTIDAD_BARRAS_DEFAULT_TXT


def _bump_cantidad_textbox(tb, delta):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if s:
            n = int(float(s.replace(u",", u".")))
        else:
            n = int(_CANTIDAD_BARRAS_DEFAULT_TXT)
    except Exception:
        n = int(_CANTIDAD_BARRAS_DEFAULT_TXT)
    n = max(
        _CANTIDAD_BARRAS_MIN,
        min(_CANTIDAD_BARRAS_MAX, n + int(delta)),
    )
    tb.Text = unicode(n)


def _normalize_estribo_separacion_textbox(tb):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if not s:
            tb.Text = _ESTRIBO_SEPARACION_DEFAULT_TXT
            return
        n = int(float(s.replace(u",", u".")))
        n = max(
            _ESTRIBO_SEPARACION_MM_MIN,
            min(_ESTRIBO_SEPARACION_MM_MAX, n),
        )
        tb.Text = unicode(n)
    except Exception:
        tb.Text = _ESTRIBO_SEPARACION_DEFAULT_TXT


def _bump_estribo_separacion_textbox(tb, delta):
    if tb is None:
        return
    try:
        s = unicode(tb.Text).strip()
        if s:
            n = int(float(s.replace(u",", u".")))
        else:
            n = int(_ESTRIBO_SEPARACION_DEFAULT_TXT)
    except Exception:
        n = int(_ESTRIBO_SEPARACION_DEFAULT_TXT)
    n = max(
        _ESTRIBO_SEPARACION_MM_MIN,
        min(_ESTRIBO_SEPARACION_MM_MAX, n + int(delta)),
    )
    tb.Text = unicode(n)


_CANTIDAD_DIGITS = u"0123456789"


def _cantidad_tb_preview_text_input(sender, e):
    try:
        t = unicode(e.Text) if e.Text else u""
        if not t:
            return
        for ch in t:
            if ch not in _CANTIDAD_DIGITS:
                e.Handled = True
                return
    except Exception:
        pass


def _cantidad_tb_pasting(sender, e):
    tb = sender
    try:
        from System.Windows import DataFormats

        raw = None
        if e.DataObject.GetDataPresent(DataFormats.UnicodeText):
            raw = e.DataObject.GetData(DataFormats.UnicodeText)
        elif e.DataObject.GetDataPresent(DataFormats.Text):
            raw = e.DataObject.GetData(DataFormats.Text)
        if raw is None:
            return
        s = unicode(raw)
        digits = u"".join(c for c in s if c in _CANTIDAD_DIGITS)
        if digits == s:
            return
        e.CancelCommand()
        try:
            sel_start = int(tb.SelectionStart)
            sel_len = int(tb.SelectionLength)
        except Exception:
            sel_start = 0
            sel_len = 0
        try:
            caret = int(tb.CaretIndex)
        except Exception:
            caret = sel_start
        t0 = unicode(tb.Text)
        if sel_len > 0:
            new_t = t0[:sel_start] + digits + t0[sel_start + sel_len :]
            tb.Text = new_t
            tb.CaretIndex = sel_start + len(digits)
        else:
            new_t = t0[:caret] + digits + t0[caret:]
            tb.Text = new_t
            tb.CaretIndex = caret + len(digits)
    except Exception:
        pass


_WINDOW_OPEN_MS = 180
_WINDOW_CLOSE_MS = 180
_WINDOW_SLIDE_PX = 10.0

# Misma plantilla visual que `enfierrado_fundacion_aislada._ENFIERRADO_FUND_XAML` (sin bloque propagación).
_ENFIERRADO_VIGAS_XAML = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Armadura vigas"
    SizeToContent="WidthAndHeight"
    MinHeight="400" MinWidth="480"
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
    <Style x:Key="LabelHint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#95B8CC"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="CapaRowLabel" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
    <Style x:Key="BtnGhost" TargetType="Button">
      <Setter Property="Background"      Value="#0E1B32"/>
      <Setter Property="Foreground"      Value="#C8E4EF"/>
      <Setter Property="FontWeight"      Value="Normal"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Padding"         Value="16,10"/>
      <Setter Property="BorderBrush"     Value="#1A3A4D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Root" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Root" Property="Background" Value="#152A45"/>
                <Setter TargetName="Root" Property="BorderBrush" Value="#5BC0DE"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Root" Property="Background" Value="#0C1628"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="ComboDiamStd" TargetType="ComboBox" BasedOn="{StaticResource Combo}">
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="HorizontalAlignment" Value="Left"/>
      <Setter Property="Width" Value="100"/>
      <Setter Property="Height" Value="30"/>
      <Setter Property="MinHeight" Value="30"/>
      <Setter Property="MaxHeight" Value="30"/>
    </Style>
    <Style x:Key="ComboTipoEstribo" TargetType="ComboBox" BasedOn="{StaticResource Combo}">
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="HorizontalAlignment" Value="Left"/>
      <Setter Property="Width" Value="118"/>
      <Setter Property="Height" Value="30"/>
      <Setter Property="MinHeight" Value="30"/>
      <Setter Property="MaxHeight" Value="30"/>
    </Style>

  </Window.Resources>
  <Border CornerRadius="10" Background="#0A1A2F" Padding="12"
          BorderBrush="#1A3A4D" BorderThickness="1"
          HorizontalAlignment="Left" ClipToBounds="True">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="16" ShadowDepth="0" Opacity="0.35"/>
    </Border.Effect>
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Border x:Name="TitleBar" Grid.Row="0" Background="#0E1B32" CornerRadius="6" Padding="10,8" Margin="0,0,0,10"
              BorderBrush="#21465C" BorderThickness="1" HorizontalAlignment="Stretch">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Image x:Name="ImgLogo" Width="44" Height="44" Grid.Column="0"
                 Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center"/>
          <StackPanel Grid.Column="1" VerticalAlignment="Center">
            <TextBlock Text="Armadura vigas" FontSize="15" FontWeight="SemiBold"
                       Foreground="#E8F4F8"/>
            <TextBlock Text="Seleccione vigas, columnas o muros y defina armadura superior (capas), empalmes e inferior"
                       FontSize="11" Foreground="#95B8CC" Margin="0,6,0,0" TextWrapping="Wrap"/>
          </StackPanel>
          <Button x:Name="BtnClose" Grid.Column="2" Style="{StaticResource BtnCloseX_MinimalNoBg}"
                  VerticalAlignment="Center" ToolTip="Cerrar"/>
        </Grid>
      </Border>

      <StackPanel x:Name="SvContenido" Grid.Row="1" Grid.IsSharedSizeScope="True"
                  VerticalAlignment="Top" HorizontalAlignment="Stretch">
          <Button x:Name="BtnSeleccionar" Content="Seleccionar en Modelo"
                  Style="{StaticResource BtnSelectOutline}"
                  HorizontalAlignment="Stretch" Margin="0,0,0,4"/>

          <StackPanel x:Name="PnlEmpalmesSup" Visibility="Collapsed" Margin="0,0,0,6">
              <TextBlock Text="Trazo inicial mayor a 12 metros" Style="{StaticResource LabelSmall}" Margin="0,0,0,2"/>
              <TextBlock x:Name="TxtEmpalmesHint" Style="{StaticResource LabelHint}"
                         Text="Seleccionar vigas que definirán empalmes. Los empalmes se generarán al centro de las vigas seleccionadas."/>
              <Button x:Name="BtnVigasEmpalme" Content="Seleccionar vigas para realizar empalmes"
                      Style="{StaticResource BtnSelectOutline}" HorizontalAlignment="Stretch" Margin="0,0,0,0" IsEnabled="False"/>
          </StackPanel>

          <StackPanel Margin="0,4,0,8" HorizontalAlignment="Stretch">
            <GroupBox Style="{StaticResource GbParams}"
                      Margin="0,0,0,8" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
            <GroupBox.Header>
              <Grid VerticalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="HdrSupNumCapasIzq"/>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="HdrSupNumCapasDer"/>
                </Grid.ColumnDefinitions>
                <CheckBox Grid.Column="0" x:Name="ChkSuperior" IsChecked="True" Content="Armadura superior"
                          Foreground="#C8E4EF" FontWeight="SemiBold" VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0,0,0">
                  <TextBlock Text="Nº de capas" Foreground="#95B8CC" FontSize="11"
                             VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <Border Width="50" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="TxtNumCapasSuperiores" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="1" TextAlignment="Center" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnNumCapasSupUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲" ToolTip="Más capas"/>
                          <RepeatButton x:Name="BtnNumCapasSupDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼" ToolTip="Menos capas"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                </StackPanel>
              </Grid>
            </GroupBox.Header>
            <StackPanel x:Name="PanelSuperior">
              <Grid Margin="0,0,0,0">
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                  <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                  <ColumnDefinition Width="8"/>
                  <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Row="0" Grid.Column="1" Text="Cantidad de barras" Style="{StaticResource LabelSmall}" VerticalAlignment="Bottom"/>
                <StackPanel Grid.Row="0" Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Bottom">
                  <TextBlock Text="ø" FontSize="12" Foreground="#5BC0DE" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,6,0"/>
                  <TextBlock Text="Diámetro (mm)" Style="{StaticResource LabelSmall}" VerticalAlignment="Bottom"/>
                </StackPanel>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="1ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                <Border Grid.Row="1" Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                        BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="20"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="CmbSupCant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                             Text="2" VerticalContentAlignment="Center"/>
                    <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                            BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                      <Grid>
                        <Grid.RowDefinitions>
                          <RowDefinition Height="*"/>
                          <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <RepeatButton x:Name="BtnCmbSupCantUp" Grid.Row="0"
                                      Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                        <RepeatButton x:Name="BtnCmbSupCantDown" Grid.Row="1"
                                      Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                      </Grid>
                    </Border>
                  </Grid>
                </Border>
                <ComboBox x:Name="CmbSupDiam" Grid.Row="1" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                          IsEditable="False" VerticalAlignment="Center">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
                <Grid x:Name="PnlCapa2" Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="4" Margin="0,2,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="2ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                  <Border Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="CmbSup2Cant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="2" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnCmbSup2CantUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                          <RepeatButton x:Name="BtnCmbSup2CantDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                  <ComboBox x:Name="CmbSup2Diam" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                            IsEditable="False" VerticalAlignment="Center">
                    <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                  </ComboBox>
                </Grid>
                <Grid x:Name="PnlCapa3" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="4" Margin="0,2,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="3ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                  <Border Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="CmbSup3Cant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="2" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnCmbSup3CantUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                          <RepeatButton x:Name="BtnCmbSup3CantDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                  <ComboBox x:Name="CmbSup3Diam" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                            IsEditable="False" VerticalAlignment="Center">
                    <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                  </ComboBox>
                </Grid>
              </Grid>
              <Grid Margin="0,8,0,4" VerticalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="HdrSupNumCapasIzq"/>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="HdrSupNumCapasDer"/>
                </Grid.ColumnDefinitions>
                <CheckBox Grid.Column="0" x:Name="ChkColocarSuple" IsChecked="True" Content="Suple superior"
                          Foreground="#C8E4EF" FontWeight="SemiBold" VerticalAlignment="Center"/>
                <StackPanel x:Name="PnlSupleSupNumCapasEdits" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0,0,0">
                  <TextBlock Text="Nº de capas" Foreground="#95B8CC" FontSize="11"
                             VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <Border Width="50" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="TxtNumCapasSupleSup" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="1" TextAlignment="Center" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnNumCapasSupleSupUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲" ToolTip="Más capas de suple"/>
                          <RepeatButton x:Name="BtnNumCapasSupleSupDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼" ToolTip="Menos capas de suple"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                </StackPanel>
              </Grid>
              <StackPanel x:Name="PnlSupleEdits">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" x:Name="TxtSupleSupCapaLbl" Text="2ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                  <Border Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="CmbSupleCant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="2" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnCmbSupleCantUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                          <RepeatButton x:Name="BtnCmbSupleCantDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                  <ComboBox x:Name="CmbSupleDiam" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                            IsEditable="False" VerticalAlignment="Center">
                    <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                  </ComboBox>
                </Grid>
                <Grid x:Name="PnlSupleCapa2" Margin="0,2,0,0" Visibility="Collapsed">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" x:Name="TxtSupleSup2CapaLbl" Text="3ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                  <Border Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="CmbSuple2Cant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="2" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnCmbSuple2CantUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                          <RepeatButton x:Name="BtnCmbSuple2CantDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                  <ComboBox x:Name="CmbSuple2Diam" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                            IsEditable="False" VerticalAlignment="Center">
                    <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                  </ComboBox>
                </Grid>
              </StackPanel>
            </StackPanel>
          </GroupBox>

            <GroupBox Style="{StaticResource GbParams}"
                      Margin="0,0,0,8" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
            <GroupBox.Header>
              <Grid VerticalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="HdrSupNumCapasIzq"/>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="HdrSupNumCapasDer"/>
                </Grid.ColumnDefinitions>
                <CheckBox Grid.Column="0" x:Name="ChkInferior" IsChecked="True" Content="Armadura inferior"
                          Foreground="#C8E4EF" FontWeight="SemiBold" VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0,0,0">
                  <TextBlock Text="Nº de capas" Foreground="#95B8CC" FontSize="11"
                             VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <Border Width="50" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="TxtNumCapasInferiores" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="1" TextAlignment="Center" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnNumCapasInfUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲" ToolTip="Más capas"/>
                          <RepeatButton x:Name="BtnNumCapasInfDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼" ToolTip="Menos capas"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                </StackPanel>
              </Grid>
            </GroupBox.Header>
            <StackPanel x:Name="PanelInferior">
              <Grid Margin="0,0,0,0">
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                  <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                  <ColumnDefinition Width="8"/>
                  <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Row="0" Grid.Column="1" Text="Cantidad de barras" Style="{StaticResource LabelSmall}" VerticalAlignment="Bottom"/>
                <StackPanel Grid.Row="0" Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Bottom">
                  <TextBlock Text="ø" FontSize="12" Foreground="#5BC0DE" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,6,0"/>
                  <TextBlock Text="Diámetro (mm)" Style="{StaticResource LabelSmall}" VerticalAlignment="Bottom"/>
                </StackPanel>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="1ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                <Border Grid.Row="1" Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                        BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="20"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="CmbInfCant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                             Text="2" VerticalContentAlignment="Center"/>
                    <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                            BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                      <Grid>
                        <Grid.RowDefinitions>
                          <RowDefinition Height="*"/>
                          <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <RepeatButton x:Name="BtnCmbInfCantUp" Grid.Row="0"
                                      Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                        <RepeatButton x:Name="BtnCmbInfCantDown" Grid.Row="1"
                                      Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                      </Grid>
                    </Border>
                  </Grid>
                </Border>
                <ComboBox x:Name="CmbInfDiam" Grid.Row="1" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                          IsEditable="False" VerticalAlignment="Center">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
                <Grid x:Name="PnlInfCapa2" Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="4" Margin="0,2,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="2ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                  <Border Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="CmbInf2Cant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="2" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnCmbInf2CantUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                          <RepeatButton x:Name="BtnCmbInf2CantDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                  <ComboBox x:Name="CmbInf2Diam" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                            IsEditable="False" VerticalAlignment="Center">
                    <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                  </ComboBox>
                </Grid>
                <Grid x:Name="PnlInfCapa3" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="4" Margin="0,2,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" Text="3ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                  <Border Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="CmbInf3Cant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="2" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnCmbInf3CantUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                          <RepeatButton x:Name="BtnCmbInf3CantDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                  <ComboBox x:Name="CmbInf3Diam" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                            IsEditable="False" VerticalAlignment="Center">
                    <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                  </ComboBox>
                </Grid>
              </Grid>
              <CheckBox x:Name="ChkColocarSupleInf" IsChecked="True" Content="Suple inferior"
                        Foreground="#C8E4EF" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,8,0,4"/>
              <StackPanel x:Name="PnlSupleInfEdits">
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0" x:Name="TxtSupleInfCapaLbl" Text="2ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                  <Border Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                          BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="CmbSupleInfCant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                               Text="2" VerticalContentAlignment="Center"/>
                      <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                              BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                        <Grid>
                          <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                          </Grid.RowDefinitions>
                          <RepeatButton x:Name="BtnCmbSupleInfCantUp" Grid.Row="0"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                          <RepeatButton x:Name="BtnCmbSupleInfCantDown" Grid.Row="1"
                                        Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                        </Grid>
                      </Border>
                    </Grid>
                  </Border>
                  <ComboBox x:Name="CmbSupleInfDiam" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                            IsEditable="False" VerticalAlignment="Center">
                    <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                  </ComboBox>
                </Grid>
              </StackPanel>
            </StackPanel>
          </GroupBox>

            <GroupBox Style="{StaticResource GbParams}" Margin="0,0,0,8" HorizontalAlignment="Stretch">
            <GroupBox.Header>
              <Grid VerticalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="HdrSupNumCapasIzq"/>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="HdrSupNumCapasDer"/>
                </Grid.ColumnDefinitions>
                <CheckBox Grid.Column="0" x:Name="ChkEstribos" IsChecked="True" Content="Estribos"
                          Foreground="#C8E4EF" FontWeight="SemiBold" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Text="" Margin="10,0,0,0"/>
              </Grid>
            </GroupBox.Header>
            <StackPanel x:Name="PnlEstribosParams" IsEnabled="True" VerticalAlignment="Top">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                  <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                  <ColumnDefinition Width="8"/>
                  <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Bottom">
                  <TextBlock Text="ø" FontSize="12" Foreground="#5BC0DE" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,6,0"/>
                  <TextBlock Text="Diámetro (mm)" Style="{StaticResource LabelSmall}" VerticalAlignment="Bottom"/>
                </StackPanel>
                <TextBlock Grid.Row="0" Grid.Column="3" Text="Separación (mm)" Style="{StaticResource LabelSmall}" VerticalAlignment="Bottom"/>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="Ext." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                <ComboBox x:Name="CmbEstriboExtDiam" Grid.Row="1" Grid.Column="1" Style="{StaticResource ComboDiamStd}"
                          IsEditable="False" VerticalAlignment="Center">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
                <Border Grid.Row="1" Grid.Column="3" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                        BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="20"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtEstriboExtSep" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                             Text="200" VerticalContentAlignment="Center"/>
                    <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                            BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                      <Grid>
                        <Grid.RowDefinitions>
                          <RowDefinition Height="*"/>
                          <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <RepeatButton x:Name="BtnEstriboExtSepUp" Grid.Row="0" Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                        <RepeatButton x:Name="BtnEstriboExtSepDown" Grid.Row="1" Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                      </Grid>
                    </Border>
                  </Grid>
                </Border>
                <TextBlock Grid.Row="2" Grid.Column="0" Text="Cent." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center" Margin="0,2,0,0"/>
                <ComboBox x:Name="CmbEstriboCentDiam" Grid.Row="2" Grid.Column="1" Style="{StaticResource ComboDiamStd}"
                          IsEditable="False" VerticalAlignment="Center" Margin="0,2,0,0">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
                <Border Grid.Row="2" Grid.Column="3" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                        BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center" Margin="0,2,0,0">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="20"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtEstriboCentSep" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                             Text="200" VerticalContentAlignment="Center"/>
                    <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                            BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                      <Grid>
                        <Grid.RowDefinitions>
                          <RowDefinition Height="*"/>
                          <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <RepeatButton x:Name="BtnEstriboCentSepUp" Grid.Row="0" Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                        <RepeatButton x:Name="BtnEstriboCentSepDown" Grid.Row="1" Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                      </Grid>
                    </Border>
                  </Grid>
                </Border>
                <Border Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="4" Height="1"
                        Background="#1A3A4D" Margin="0,12,0,0" VerticalAlignment="Center" SnapsToDevicePixels="True"/>
                <StackPanel Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="3" Orientation="Horizontal"
                            VerticalAlignment="Bottom" Margin="0,8,0,4">
                  <TextBlock Text="Tipo" FontSize="12" Foreground="#5BC0DE" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <TextBlock Text="de estribo" Style="{StaticResource LabelSmall}" VerticalAlignment="Bottom"/>
                </StackPanel>
                <ComboBox x:Name="CmbEstriboTipo" Grid.Row="5" Grid.Column="1" Grid.ColumnSpan="3"
                          Style="{StaticResource ComboTipoEstribo}"
                          IsEditable="False" VerticalAlignment="Center" HorizontalAlignment="Left"
                          Width="208" Margin="0,0,0,0">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
              </Grid>
            </StackPanel>
          </GroupBox>

            <GroupBox Style="{StaticResource GbParams}"
                      Margin="0,0,0,0" HorizontalAlignment="Stretch">
            <GroupBox.Header>
              <CheckBox x:Name="ChkLaterales" IsChecked="True" Content="Barras laterales"
                        Foreground="#C8E4EF" FontWeight="SemiBold" VerticalAlignment="Center"/>
            </GroupBox.Header>
            <StackPanel x:Name="PnlLateralesParams" IsEnabled="True" VerticalAlignment="Top">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto" SharedSizeGroup="CapaLbl" MinWidth="40"/>
                  <ColumnDefinition Width="100" SharedSizeGroup="SpinCant"/>
                  <ColumnDefinition Width="8"/>
                  <ColumnDefinition Width="100" SharedSizeGroup="SpinDiam"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Row="0" Grid.Column="1" Text="Cantidad de barras" Style="{StaticResource LabelSmall}" VerticalAlignment="Bottom"/>
                <StackPanel Grid.Row="0" Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Bottom">
                  <TextBlock Text="ø" FontSize="12" Foreground="#5BC0DE" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,6,0"/>
                  <TextBlock Text="Diámetro (mm)" Style="{StaticResource LabelSmall}" VerticalAlignment="Bottom"/>
                </StackPanel>
                <TextBlock Grid.Row="1" Grid.Column="0" Text="1ªC." Style="{StaticResource CapaRowLabel}" VerticalAlignment="Center"/>
                <Border Grid.Row="1" Grid.Column="1" Width="100" Height="30" CornerRadius="4" Background="#050E18" HorizontalAlignment="Left"
                        BorderBrush="#1A3A4D" BorderThickness="1" SnapsToDevicePixels="True" VerticalAlignment="Center">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="20"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="CmbLatCant" Grid.Column="0" Style="{StaticResource CantSpinnerText}"
                             Text="2" VerticalContentAlignment="Center"/>
                    <Border Grid.Column="1" Background="#0E1B32" BorderBrush="#1A3A4D"
                            BorderThickness="1,0,0,0" CornerRadius="0,4,4,0" ClipToBounds="True">
                      <Grid>
                        <Grid.RowDefinitions>
                          <RowDefinition Height="*"/>
                          <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <RepeatButton x:Name="BtnCmbLatCantUp" Grid.Row="0"
                                      Style="{StaticResource SpinRepeatBtn}" Content="▲"/>
                        <RepeatButton x:Name="BtnCmbLatCantDown" Grid.Row="1"
                                      Style="{StaticResource SpinRepeatBtn}" Content="▼"/>
                      </Grid>
                    </Border>
                  </Grid>
                </Border>
                <ComboBox x:Name="CmbLatDiam" Grid.Row="1" Grid.Column="3" Style="{StaticResource ComboDiamStd}"
                          IsEditable="False" VerticalAlignment="Center">
                  <ComboBox.ItemContainerStyle><Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/></ComboBox.ItemContainerStyle>
                </ComboBox>
              </Grid>
            </StackPanel>
          </GroupBox>

          </StackPanel>
        </StackPanel>

      <StackPanel Grid.Row="2" Margin="0,20,0,0">
        <TextBlock x:Name="TxtEstado" Text="" Foreground="#5BC0DE" FontSize="11"
                   Margin="0,0,0,8" TextWrapping="Wrap"/>
        <CheckBox x:Name="ChkModelLines" IsChecked="True" Foreground="#C8E4EF" FontSize="11"
                  Content="Model lines (guías de trazo en el modelo)" Margin="0,0,0,8"/>
        <Button x:Name="BtnColocar" Content="Colocar armadura"
                Style="{StaticResource BtnPrimary}"
                HorizontalAlignment="Stretch"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
"""


class EnfierradoVigasSelectionFilter(ISelectionFilter):
    """Structural Framing, Structural Columns y Walls (por categoría)."""

    def AllowElement(self, elem):
        try:
            if elem is None:
                return False
            cat = elem.Category
            if cat is None:
                return False
            return element_id_to_int(cat.Id) in _ALLOWED_SELECTION_CAT_IDS
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


class SeleccionarVigasHandler(IExternalEventHandler):
    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        from Autodesk.Revit.UI.Selection import ObjectType

        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win._set_estado(u"No hay documento activo.")
            return
        doc = uidoc.Document
        flt = EnfierradoVigasSelectionFilter()
        try:
            refs = list(
                uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    flt,
                    u"Seleccione vigas, columnas o muros (Structural Framing, Columns, Walls). Finalice con Finalizar.",
                )
            )
        except Exception:
            refs = []
            win._set_estado(u"Selección cancelada.")
            try:
                win._show_with_fade()
            except Exception:
                pass
            return

        if not refs:
            win._set_estado(u"Sin elementos.")
            try:
                win._show_with_fade()
            except Exception:
                pass
            return

        ids = []
        for r in refs:
            try:
                ids.append(r.ElementId)
            except Exception:
                pass
        win._document = doc
        win._selected_element_ids = ids
        win._refresh_selection_text()
        try:
            win._refresh_empalmes_panel_from_selection()
        except Exception:
            pass
        try:
            win._refresh_laterales_cantidad_desde_seleccion()
        except Exception:
            pass
        win._set_estado(
            u"{0} elemento(s) seleccionado(s).".format(len(ids))
        )
        try:
            win._show_with_fade()
        except Exception:
            pass

    def GetName(self):
        return u"SeleccionarVigasEnfierrado"


class EmpalmeFramingOnlyFilter(ISelectionFilter):
    """Solo instancias Structural Framing."""

    def AllowElement(self, elem):
        try:
            if elem is None or elem.Category is None:
                return False
            return element_id_to_int(elem.Category.Id) == int(
                BuiltInCategory.OST_StructuralFraming
            )
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


class PickEmpalmeFramingHandler(IExternalEventHandler):
    """Selección de vigas para troceo / empalmes (solo Structural Framing)."""

    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        from Autodesk.Revit.UI.Selection import ObjectType

        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            win._set_estado(u"No hay documento activo.")
            return
        doc = uidoc.Document
        flt = EmpalmeFramingOnlyFilter()
        try:
            refs = list(
                uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    flt,
                    u"Seleccione vigas Structural Framing para empalmes (plano al centro del eje). Finalice con Finalizar.",
                )
            )
        except Exception:
            refs = []
            win._set_estado(u"Selección de vigas empalme cancelada.")
            try:
                win._show_with_fade()
            except Exception:
                pass
            return
        if not refs:
            win._set_estado(u"Sin vigas de empalme.")
            try:
                win._show_with_fade()
            except Exception:
                pass
            return
        ids = []
        for r in refs:
            try:
                ids.append(r.ElementId)
            except Exception:
                pass
        win._empalme_framing_ids = ids
        win._document = doc
        win._refresh_empalme_text()
        win._set_estado(u"{0} viga(s) de troceo por empalme.".format(len(ids)))
        try:
            win._show_with_fade()
        except Exception:
            pass

    def GetName(self):
        return u"PickEmpalmeFramingVigas"


def _parse_cantidad_capa_ventana(enfierrado_win, combo_xname, forzar_combo_habilitado=False):
    if enfierrado_win is None:
        return 1
    cmb = enfierrado_win._win.FindName(combo_xname)
    if cmb is None:
        return 1
    prev_en = True
    if forzar_combo_habilitado:
        try:
            prev_en = cmb.IsEnabled
        except Exception:
            prev_en = True
        try:
            cmb.IsEnabled = True
            cmb.UpdateLayout()
        except Exception:
            pass
    try:
        try:
            t = cmb.Text
        except Exception:
            t = None
        if not t:
            try:
                t = cmb.SelectedItem
            except Exception:
                t = None
        n = int(float(unicode(t).replace(u",", u".")))
        return max(
            _CANTIDAD_BARRAS_MIN,
            min(_CANTIDAD_BARRAS_MAX, n),
        )
    except Exception:
        return 1
    finally:
        if forzar_combo_habilitado:
            try:
                cmb.IsEnabled = prev_en
            except Exception:
                pass


def _parse_cantidad_superior_ventana(enfierrado_win):
    return _parse_cantidad_capa_ventana(enfierrado_win, "CmbSupCant")


def _parse_cantidad_suple_ventana(enfierrado_win):
    return _parse_cantidad_capa_ventana(enfierrado_win, "CmbSupleCant")


def _parse_capas_superiores_ventana(enfierrado_win):
    """Número de capas de armadura superior longitudinal (1–3), desde el control incremental."""
    if enfierrado_win is None:
        return 1
    tb = enfierrado_win._win.FindName("TxtNumCapasSuperiores")
    if tb is not None:
        try:
            s = unicode(tb.Text).strip()
            if s:
                n = int(float(s.replace(u",", u".")))
                return max(
                    _CAPAS_ARMADURA_MIN,
                    min(_CAPAS_ARMADURA_MAX, n),
                )
        except Exception:
            pass
    return int(_CAPAS_DEFAULT_TXT)


def _parse_capas_inferiores_ventana(enfierrado_win):
    """Número de capas de armadura inferior longitudinal (1–3)."""
    if enfierrado_win is None:
        return 1
    tb = enfierrado_win._win.FindName("TxtNumCapasInferiores")
    if tb is not None:
        try:
            s = unicode(tb.Text).strip()
            if s:
                n = int(float(s.replace(u",", u".")))
                return max(
                    _CAPAS_ARMADURA_MIN,
                    min(_CAPAS_ARMADURA_MAX, n),
                )
        except Exception:
            pass
    return int(_CAPAS_DEFAULT_TXT)


def _parse_capas_suple_superior_ventana(enfierrado_win):
    """Capas de suple superior (1–2), desde ``TxtNumCapasSupleSup``."""
    if enfierrado_win is None:
        return 1
    tb = enfierrado_win._win.FindName("TxtNumCapasSupleSup")
    if tb is not None:
        try:
            s = unicode(tb.Text).strip()
            if s:
                n = int(float(s.replace(u",", u".")))
                return max(
                    _CAPAS_SUPLE_MIN,
                    min(_CAPAS_SUPLE_MAX, n),
                )
        except Exception:
            pass
    return int(_CAPAS_SUPLE_DEFAULT_TXT)


def _suple_capa_label_text(n_capas):
    """Texto «(N+1)ªC.» para fila suple según N capas longitudinales activas (1–3)."""
    n = max(_CAPAS_ARMADURA_MIN, min(_CAPAS_ARMADURA_MAX, int(n_capas or 1)))
    return u"{0}ªC.".format(n + 1)


def _suple_sup_layer_label(n_main, idx_suple):
    """Etiqueta de fila suple superior: ``(N_main+1+k)ªC.`` con k = 0 ó 1."""
    nm = max(_CAPAS_ARMADURA_MIN, min(_CAPAS_ARMADURA_MAX, int(n_main or 1)))
    k = max(0, min(1, int(idx_suple or 0)))
    return u"{0}ªC.".format(nm + 1 + k)


def _unwrap_combo_item_content(raw):
    if raw is None:
        return None
    try:
        from System.Windows.Controls import ComboBoxItem

        if isinstance(raw, ComboBoxItem):
            return raw.Content
    except Exception:
        pass
    return raw


def _etiqueta_diam_normalizada(s):
    if s is None:
        return u""
    try:
        t = unicode(s).strip().lower()
    except Exception:
        return u""
    for ch in (u"ø", u"Ø", u"Φ", u"∅"):
        t = t.replace(ch, u"")
    t = t.replace(u" ", u"").replace(u"mm", u"")
    return t


def _rebar_bar_type_desde_entries_row(enfierrado_win, entry):
    """
    Convierte una fila ``(RebarBarType|None, etiqueta)`` de ``_entries`` en ``RebarBarType``.
    Hay filas sintéticas ``(None, «ø12 mm»)`` sin elemento; deben resolverse al nominal
    en el documento — si no, el matcher por texto devolvía ``None`` y la geometría
    repetía el tipo de la 1.ª capa aunque la cantidad por capa fuera correcta.
    """
    if not entry:
        return None
    try:
        bt = entry[0]
        lbl = entry[1]
    except Exception:
        return None
    if bt is not None:
        return bt
    doc = getattr(enfierrado_win, "_document", None)
    if doc is None:
        return None
    try:
        from enfierrado_shaft_hashtag import resolver_bar_type_por_diametro_mm

        m = re.search(r"(\d+(?:[.,]\d+)?)", unicode(lbl or u""))
        if not m:
            return None
        mm = float(m.group(1).replace(u",", u"."))
        btr, _ex, _d = resolver_bar_type_por_diametro_mm(doc, mm)
        return btr
    except Exception:
        return None


def _rebar_bar_type_match_por_texto(entries, raw_ui):
    """
    Alinea el texto mostrado en el combo (ø12 mm, SelectionBoxItem, etc.)
    con ``(RebarBarType, etiqueta)`` de ``_entries``.
    """
    if not entries or raw_ui is None:
        return None
    raw_ui = _unwrap_combo_item_content(raw_ui)
    if raw_ui is None:
        return None
    try:
        s0 = unicode(raw_ui).strip()
    except Exception:
        return None
    if not s0:
        return None
    for bt, lbl in entries:
        try:
            if unicode(lbl).strip() == s0:
                if bt is not None:
                    return bt
        except Exception:
            continue
    n0 = _etiqueta_diam_normalizada(s0)
    if n0:
        for bt, lbl in entries:
            try:
                if _etiqueta_diam_normalizada(lbl) == n0:
                    if bt is not None:
                        return bt
            except Exception:
                continue
    try:
        m = re.search(r"(\d+(?:[.,]\d+)?)", s0)
        if m:
            val = float(m.group(1).replace(u",", u"."))
            candidatos = []
            for bt, lbl in entries:
                if bt is None:
                    continue
                try:
                    dmm = _rebar_nominal_diameter_mm(bt)
                except Exception:
                    dmm = None
                if dmm is not None and abs(float(dmm) - val) < 0.05:
                    candidatos.append(bt)
            if len(candidatos) == 1:
                return candidatos[0]
    except Exception:
        pass
    return None


def _superior_rebar_types_y_cantidades_por_capa(enfierrado_win, n_capas):
    """
    Lee cantidad y ``RebarBarType`` **solo** de los controles de cada capa activa en UI:
    1.ª → ``CmbSupCant`` / ``CmbSupDiam``; 2.ª → ``CmbSup2*``; 3.ª → ``CmbSup3*``.
    Debe llamarse tras ``_preparar_lectura_capas_superiores`` para alinear paneles con N.
    """
    n = max(1, min(3, int(n_capas or 1)))
    pairs = (
        (u"CmbSupCant", u"CmbSupDiam"),
        (u"CmbSup2Cant", u"CmbSup2Diam"),
        (u"CmbSup3Cant", u"CmbSup3Diam"),
    )
    types = []
    cants = []
    for i in range(n):
        c_nm, d_nm = pairs[i]
        # Combos de capas 2–3 pueden estar deshabilitados si el panel estuvo oculto;
        # sin habilitar al leer, WPF a veces devuelve el mismo Ø que la 1.ª capa.
        cants.append(
            _parse_cantidad_capa_ventana(
                enfierrado_win, c_nm, forzar_combo_habilitado=True
            )
        )
        types.append(
            _rebar_bar_type_desde_combo_diam(
                enfierrado_win, d_nm, forzar_combo_habilitado=True
            )
        )
    return types, cants


def _inferior_rebar_types_y_cantidades_por_capa(enfierrado_win, n_capas):
    """
    Igual que :func:`_superior_rebar_types_y_cantidades_por_capa` para la cara inferior.
    """
    n = max(1, min(3, int(n_capas or 1)))
    pairs = (
        (u"CmbInfCant", u"CmbInfDiam"),
        (u"CmbInf2Cant", u"CmbInf2Diam"),
        (u"CmbInf3Cant", u"CmbInf3Diam"),
    )
    types = []
    cants = []
    for i in range(n):
        c_nm, d_nm = pairs[i]
        cants.append(
            _parse_cantidad_capa_ventana(
                enfierrado_win, c_nm, forzar_combo_habilitado=True
            )
        )
        types.append(
            _rebar_bar_type_desde_combo_diam(
                enfierrado_win, d_nm, forzar_combo_habilitado=True
            )
        )
    return types, cants


def _parse_cantidad_suple_inferior_ventana(enfierrado_win):
    return _parse_cantidad_capa_ventana(enfierrado_win, "CmbSupleInfCant")


def _parse_cantidad_laterales_ventana(enfierrado_win):
    return _parse_cantidad_capa_ventana(enfierrado_win, "CmbLatCant")


def _rebar_bar_type_suple_inferior_desde_ventana(enfierrado_win):
    return _rebar_bar_type_desde_combo_diam(enfierrado_win, "CmbSupleInfDiam")


def _rebar_bar_type_desde_combo_diam(
    enfierrado_win, combo_xname, forzar_combo_habilitado=False
):
    """
    Resuelve ``RebarBarType`` desde un ``ComboBox`` de diámetros rellenado en
    ``_cargar_combos_diametro`` (índice + texto del ítem por si SelectedIndex falla en WPF).
    """
    if enfierrado_win is None:
        return None
    entries = getattr(enfierrado_win, "_entries", None)
    if not entries:
        return None
    cmb = enfierrado_win._win.FindName(combo_xname)
    if cmb is None:
        return None
    prev_en = True
    if forzar_combo_habilitado:
        try:
            prev_en = cmb.IsEnabled
        except Exception:
            prev_en = True
        try:
            cmb.IsEnabled = True
            cmb.UpdateLayout()
        except Exception:
            pass
    try:
        nitems = int(cmb.Items.Count)
        idx = int(cmb.SelectedIndex)
        # 0) Índice → fila de ``_entries`` (misma orden que ``Items`` en ``_cargar_combos_diametro``),
        #    incluyendo resolución de filas (None, etiqueta) vía nominal en documento.
        if nitems > 0 and 0 <= idx < nitems and 0 <= idx < len(entries):
            bt = _rebar_bar_type_desde_entries_row(enfierrado_win, entries[idx])
            if bt is not None:
                return bt
        # 1) Contenido del ítem en esa posición; SelectionBoxItem puede ir desfasado en WPF.
        if nitems > 0 and 0 <= idx < nitems:
            try:
                it = cmb.Items[idx]
            except Exception:
                it = None
            bt = _rebar_bar_type_match_por_texto(entries, it)
            if bt is not None:
                return bt
        # 2) Texto pintado en la caja cerrada (respaldo si idx no sirve)
        for _prop in (u"SelectionBoxItem",):
            try:
                raw = getattr(cmb, _prop, None)
            except Exception:
                raw = None
            bt = _rebar_bar_type_match_por_texto(entries, raw)
            if bt is not None:
                return bt
        for attr in ("SelectedItem", "Text"):
            try:
                raw = getattr(cmb, attr, None)
            except Exception:
                raw = None
            bt = _rebar_bar_type_match_por_texto(entries, raw)
            if bt is not None:
                return bt
        return None
    finally:
        if forzar_combo_habilitado:
            try:
                cmb.IsEnabled = prev_en
            except Exception:
                pass


def _rebar_bar_type_superior_desde_ventana(enfierrado_win):
    """``RebarBarType`` del combo armadura superior, o ``None``."""
    return _rebar_bar_type_desde_combo_diam(enfierrado_win, "CmbSupDiam")


def _rebar_bar_type_suple_desde_ventana(enfierrado_win):
    """``RebarBarType`` del combo diámetro suple superior (1.ª fila), o ``None``."""
    return _rebar_bar_type_desde_combo_diam(enfierrado_win, "CmbSupleDiam")


def _superior_suple_tipos_y_cantidades(enfierrado_win, n_capas_suple):
    """
    Listas alineadas con ``n_capas_suple`` (1–2): cantidad y tipo por fila suple superior.
    """
    n = max(
        _CAPAS_SUPLE_MIN,
        min(_CAPAS_SUPLE_MAX, int(n_capas_suple or 1)),
    )
    pairs = (
        (u"CmbSupleCant", u"CmbSupleDiam"),
        (u"CmbSuple2Cant", u"CmbSuple2Diam"),
    )
    types = []
    cants = []
    for i in range(n):
        c_nm, d_nm = pairs[i]
        cants.append(
            _parse_cantidad_capa_ventana(
                enfierrado_win, c_nm, forzar_combo_habilitado=True
            )
        )
        types.append(
            _rebar_bar_type_desde_combo_diam(
                enfierrado_win, d_nm, forzar_combo_habilitado=True
            )
        )
    return types, cants


class ColocarArmaduraVigasStubHandler(IExternalEventHandler):
    """Incluye desplaz. eje V (ancho/2−25 mm), estiram./troceo y recorte en puntas; ModelCurve(s)."""

    def __init__(self, window_ref):
        self._window_ref = window_ref

    def Execute(self, uiapp):
        win = self._window_ref()
        if win is None:
            return
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            return
        doc = uidoc.Document
        view = uidoc.ActiveView

        from geometria_viga_cara_superior_detalle import (
            crear_detail_lines_largo_cara_superior_en_vista,
            filtrar_obstaculos_seleccion_no_framing,
            filtrar_solo_structural_framing,
        )

        ids_sel = win._selected_element_ids or []
        framing = filtrar_solo_structural_framing(doc, ids_sel)
        obstaculos = filtrar_obstaculos_seleccion_no_framing(doc, ids_sel)
        bloques = []
        try:
            bloques.append(win._texto_task_dialog_colocar_armadura())
        except Exception as ex:
            try:
                bloques.append(u"(Resumen UI: {0})".format(unicode(ex)))
            except Exception:
                bloques.append(u"(Resumen UI no disponible.)")
        bloques.append(u"")
        bloques.append(u"— Structural Framing —")
        if not framing:
            bloques.append(
                u"Ningún elemento seleccionado es Structural Framing; no se dibujaron líneas."
            )
            n_creadas = 0
            n_rebar = 0
            avisos = []
        else:
            bloques.append(
                u"Instancias de armadura estructural (vigas) usadas: {0}.".format(len(framing))
            )
            chk_sup = win._win.FindName("ChkSuperior")
            sup_on = chk_sup is None or chk_sup.IsChecked == True
            chk_suple = win._win.FindName("ChkColocarSuple")
            suple_on = (
                sup_on
                and (chk_suple is None or chk_suple.IsChecked == True)
            )
            n_capas_sup = _parse_capas_superiores_ventana(win) if sup_on else 1
            if doc is not None and not getattr(win, "_entries", None):
                try:
                    win._document = doc
                    win._cargar_combos_diametro()
                except Exception:
                    pass
            tipos_cap = None
            cants_cap = None
            if sup_on:
                try:
                    win._preparar_lectura_capas_superiores()
                except Exception:
                    pass
                n_capas_sup = _parse_capas_superiores_ventana(win)
                tipos_cap, cants_cap = _superior_rebar_types_y_cantidades_por_capa(
                    win, n_capas_sup
                )
            bar_tipo = tipos_cap[0] if sup_on and tipos_cap else None
            bar_suple = _rebar_bar_type_suple_desde_ventana(win) if suple_on else None
            if suple_on and bar_suple is None and bar_tipo is not None:
                bar_suple = bar_tipo
                bloques.append(
                    u"Aviso: suple — no se resolvió el diámetro en el combo; "
                    u"se usa el tipo de la primera capa superior (ajuste «Suple superior» si corresponde)."
                )
                bloques.append(u"")
            elif suple_on and bar_suple is None:
                bloques.append(
                    u"Aviso: suple activo pero no se resolvió el tipo de barra en "
                    u"«Diámetro» (recargue el diálogo o elija otro diámetro)."
                )
                bloques.append(u"")
            n_capas_suple_sup = 1
            tipos_sl = None
            cants_sl = None
            if suple_on:
                try:
                    win._sync_suple_capa_labels()
                except Exception:
                    pass
                n_capas_suple_sup = _parse_capas_suple_superior_ventana(win)
                tipos_sl, cants_sl = _superior_suple_tipos_y_cantidades(
                    win, n_capas_suple_sup
                )
                for j in range(len(tipos_sl or [])):
                    if tipos_sl[j] is None:
                        tipos_sl[j] = bar_suple if bar_suple is not None else bar_tipo
                if tipos_sl and tipos_sl[0] is not None:
                    bar_suple = tipos_sl[0]
            chk_inf = win._win.FindName("ChkInferior")
            inf_on = chk_inf is None or chk_inf.IsChecked == True
            troceo_on = bool(sup_on or inf_on)
            n_creadas = 0
            n_rebar = 0
            avisos = []

            inf_ejecutar = False
            chk_suple_inf = None
            suple_inf_on = False
            n_capas_inf = 1
            tipos_inf = None
            cants_inf = None
            bar_inf = None
            bar_suple_inf = None

            if inf_on:
                chk_suple_inf = win._win.FindName("ChkColocarSupleInf")
                suple_inf_on = (
                    inf_on
                    and (
                        chk_suple_inf is None
                        or chk_suple_inf.IsChecked == True
                    )
                )
                try:
                    win._preparar_lectura_capas_inferiores()
                except Exception:
                    pass
                n_capas_inf = _parse_capas_inferiores_ventana(win)
                tipos_inf, cants_inf = _inferior_rebar_types_y_cantidades_por_capa(
                    win, n_capas_inf
                )
                bar_inf = tipos_inf[0] if tipos_inf else None
                bar_suple_inf = (
                    _rebar_bar_type_suple_inferior_desde_ventana(win)
                    if suple_inf_on
                    else None
                )
                if suple_inf_on and bar_suple_inf is None and bar_inf is not None:
                    bar_suple_inf = bar_inf
                    bloques.append(
                        u"Aviso: suple inferior — no se resolvió el diámetro en el combo; "
                        u"se usa el tipo de la primera capa inferior."
                    )
                    bloques.append(u"")
                elif suple_inf_on and bar_suple_inf is None:
                    bloques.append(
                        u"Aviso: suple inferior activo pero sin RebarBarType en "
                        u"«Diámetro»; revise el combo."
                    )
                    bloques.append(u"")
                if bar_inf is None:
                    bloques.append(
                        u"Cara inferior: no se resolvió RebarBarType en la 1.ª capa; "
                        u"no se creó armadura inferior."
                    )
                else:
                    inf_ejecutar = True

            if sup_on or inf_ejecutar:
                from Autodesk.Revit.DB import Transaction

                chk_lat_arm = win._win.FindName("ChkLaterales")
                lat_arm_on = (
                    chk_lat_arm is not None and chk_lat_arm.IsChecked == True
                )
                bar_lat_tipo = (
                    _rebar_bar_type_desde_combo_diam(win, "CmbLatDiam")
                    if lat_arm_on
                    else None
                )
                n_lat_arm = (
                    _parse_cantidad_laterales_ventana(win)
                    if lat_arm_on
                    else 1
                )
                crear_lat_sup = bool(
                    lat_arm_on
                    and sup_on
                    and bar_lat_tipo is not None
                )
                crear_lat_inf = bool(
                    lat_arm_on
                    and inf_ejecutar
                    and bar_lat_tipo is not None
                )
                if (
                    lat_arm_on
                    and bar_lat_tipo is None
                    and (sup_on or inf_ejecutar)
                ):
                    bloques.append(
                        u"Aviso: «Barras laterales» activo pero sin RebarBarType "
                        u"(diámetro); no se crean laterales."
                    )
                    bloques.append(u"")

                chk_ml = win._win.FindName("ChkModelLines")
                crear_ml = chk_ml is None or chk_ml.IsChecked == True

                t_arm = Transaction(
                    doc, u"BIMTools — Enfierrado vigas (detalle)"
                )
                t_arm.Start()
                try:
                    bloque_sup_txt = None
                    bloque_inf_txt = None
                    if sup_on:
                        nc, nr, av = crear_detail_lines_largo_cara_superior_en_vista(
                            doc,
                            view,
                            framing,
                            elementos_obstaculos=obstaculos,
                            rebar_bar_type=bar_tipo,
                            rebar_cantidad=(
                                cants_cap[0] if cants_cap else 1
                            ),
                            framing_empalme_element_ids=list(
                                win._empalme_framing_ids or []
                            ),
                            aplicar_troceo_empalmes_framing=troceo_on,
                            n_capas_superiores=int(n_capas_sup),
                            rebar_bar_types_capas=tipos_cap,
                            rebar_cantidades_capas=cants_cap,
                            rebar_bar_type_suple=bar_suple,
                            rebar_cantidad_suple=(
                                (cants_sl[0] if cants_sl else 1)
                                if suple_on
                                else 1
                            ),
                            crear_armadura_suple=bool(suple_on),
                            n_capas_suple=(
                                int(n_capas_suple_sup) if suple_on else 1
                            ),
                            rebar_bar_types_suple=(
                                tipos_sl if suple_on else None
                            ),
                            rebar_cantidades_suple=(
                                cants_sl if suple_on else None
                            ),
                            es_cara_inferior=False,
                            gestionar_transaccion=False,
                            crear_laterales_cara_superior=crear_lat_sup,
                            laterales_rebar_bar_type=bar_lat_tipo,
                            laterales_cantidad=n_lat_arm,
                            crear_model_lines=crear_ml,
                        )
                        n_creadas += nc
                        n_rebar += nr
                        avisos.extend(av or [])
                        bloque_sup_txt = (
                            u"Cara superior — Model lines: {0}. Rebar: {1}. "
                            u"(unif., −25 n, V b/2−25, +2 m/ext., −25 fin".format(
                                nc, nr
                            )
                            + (u"; troceo obst." if obstaculos else u"")
                            + u")"
                        )
                    if inf_ejecutar:
                        nci, nri, avi = (
                            crear_detail_lines_largo_cara_superior_en_vista(
                                doc,
                                view,
                                framing,
                                elementos_obstaculos=obstaculos,
                                rebar_bar_type=bar_inf,
                                rebar_cantidad=(
                                    cants_inf[0] if cants_inf else 1
                                ),
                                framing_empalme_element_ids=list(
                                    win._empalme_framing_ids or []
                                ),
                                aplicar_troceo_empalmes_framing=troceo_on,
                                n_capas_superiores=int(n_capas_inf),
                                rebar_bar_types_capas=tipos_inf,
                                rebar_cantidades_capas=cants_inf,
                                rebar_bar_type_suple=bar_suple_inf,
                                rebar_cantidad_suple=_parse_cantidad_suple_inferior_ventana(
                                    win
                                )
                                if suple_inf_on
                                else 1,
                                crear_armadura_suple=bool(suple_inf_on),
                                es_cara_inferior=True,
                                gestionar_transaccion=False,
                                crear_model_lines=crear_ml,
                                crear_laterales_cara_superior=crear_lat_inf,
                                laterales_rebar_bar_type=bar_lat_tipo,
                                laterales_cantidad=n_lat_arm,
                            )
                        )
                        n_creadas += nci
                        n_rebar += nri
                        avisos.extend(avi or [])
                        bloque_inf_txt = (
                            u"Cara inferior — Model lines: {0}. Rebar: {1}. "
                            u"(proyección plano inferior, −25 mm según n exterior (hacia interior); V b/2−25; troceo/recortes".format(
                                nci, nri
                            )
                            + (u"; obst." if obstaculos else u"")
                            + u")"
                        )
                    t_arm.Commit()
                    if bloque_sup_txt is not None:
                        bloques.append(bloque_sup_txt)
                    if bloque_inf_txt is not None:
                        bloques.append(bloque_inf_txt)
                except Exception as ex:
                    try:
                        t_arm.RollBack()
                    except Exception:
                        pass
                    n_creadas = 0
                    n_rebar = 0
                    avisos = []
                    try:
                        avisos.append(
                            u"Transacción (enfierrado vigas): {0}".format(
                                unicode(ex)
                            )
                        )
                    except Exception:
                        avisos.append(
                            u"Transacción (enfierrado vigas): error; sin cambios."
                        )
            if sup_on or inf_on:
                bloques.append(
                    u"Total — Model lines: {0}. Rebar: {1}.".format(
                        n_creadas, n_rebar
                    )
                )
            if avisos:
                bloques.append(u"")
                bloques.append(u"Avisos:")
                for a in avisos[:25]:
                    bloques.append(a)
                if len(avisos) > 25:
                    bloques.append(u"… ({0} avisos más)".format(len(avisos) - 25))
        chk_lat_res = win._win.FindName("ChkLaterales")
        if chk_lat_res is not None and chk_lat_res.IsChecked == True:
            bloques.append(u"")
            bloques.append(
                u"Laterales (cara sup./inf. según lo colocado): curva guía tras última capa, "
                u"Rebar con Fixed Number según cantidad/diámetro de la tarjeta "
                u"(avisos arriba si hubo fallos)."
            )
        chk_est_res = win._win.FindName("ChkEstribos")
        if chk_est_res is not None and chk_est_res.IsChecked == True:
            bloques.append(u"")
            bloques.append(
                u"Nota: «Estribos» figura en el resumen; "
                u"la colocación en el modelo no está enlazada aún a la geometría."
            )
        msg = u"\n".join(bloques)
        try:
            if len(msg) > 16000:
                msg = msg[:15900] + u"\n\n… (mensaje truncado por longitud)."
        except Exception:
            pass
        TaskDialog.Show(u"BIMTools — Enfierrado vigas — Resumen", msg)
        try:
            win._set_estado(
                u"Model lines: {0}, Rebar: {1} (vigas: {2}).".format(
                    n_creadas, n_rebar, len(framing)
                )
            )
        except Exception:
            pass

    def GetName(self):
        return u"ColocarArmaduraVigasStub"


def _clear_appdomain_window_key():
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _get_active_window():
    try:
        win = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        _ = win.Title
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            _clear_appdomain_window_key()
            return None
    except Exception:
        _clear_appdomain_window_key()
        return None
    return win


class EnfierradoVigasWindow(object):
    def __init__(self, revit):
        self._revit = revit
        self._document = None
        self._selected_element_ids = []
        self._empalme_framing_ids = []
        self._entries = []
        self._is_closing_with_fade = False
        self._base_top = None

        from System.Windows import RoutedEventHandler
        from System.Windows.Input import ApplicationCommands, CommandBinding, Key, KeyBinding, ModifierKeys
        from System.Windows.Markup import XamlReader

        self._win = XamlReader.Parse(_ENFIERRADO_VIGAS_XAML)

        self._seleccion_handler = SeleccionarVigasHandler(weakref.ref(self))
        self._seleccion_event = ExternalEvent.Create(self._seleccion_handler)
        self._colocar_handler = ColocarArmaduraVigasStubHandler(weakref.ref(self))
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)
        self._empalme_pick_handler = PickEmpalmeFramingHandler(weakref.ref(self))
        self._empalme_pick_event = ExternalEvent.Create(self._empalme_pick_handler)

        self._setup_ui(RoutedEventHandler)
        self._wire_commands(RoutedEventHandler, ApplicationCommands, CommandBinding, KeyBinding, Key, ModifierKeys)
        self._wire_lifecycle_handlers()

    def _wire_lifecycle_handlers(self):
        try:
            from System.Windows import RoutedEventHandler

            def _on_closed(sender, args):
                _clear_appdomain_window_key()

            self._win.Closed += RoutedEventHandler(_on_closed)
        except Exception:
            pass

    def _setup_ui(self, RoutedEventHandler):
        from System.IO import FileAccess, FileMode, FileStream
        from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage

        try:
            img = self._win.FindName("ImgLogo")
            if img is not None:
                for logo_path in get_logo_paths():
                    if os.path.isfile(logo_path):
                        stream = None
                        try:
                            stream = FileStream(logo_path, FileMode.Open, FileAccess.Read)
                            bmp = BitmapImage()
                            bmp.BeginInit()
                            bmp.StreamSource = stream
                            bmp.CacheOption = BitmapCacheOption.OnLoad
                            bmp.EndInit()
                            bmp.Freeze()
                            img.Source = bmp
                        finally:
                            if stream is not None:
                                try:
                                    stream.Dispose()
                                except Exception:
                                    pass
                        break
        except Exception:
            pass

        btn_sel = self._win.FindName("BtnSeleccionar")
        if btn_sel is not None:
            btn_sel.Click += RoutedEventHandler(self._on_seleccionar)
        btn_close = self._win.FindName("BtnClose")
        if btn_close is not None:
            btn_close.Click += RoutedEventHandler(lambda s, e: self._close_with_fade())
        btn_col = self._win.FindName("BtnColocar")
        if btn_col is not None:
            btn_col.Click += RoutedEventHandler(self._on_colocar)
        btn_emp = self._win.FindName("BtnVigasEmpalme")
        if btn_emp is not None:
            btn_emp.Click += RoutedEventHandler(self._on_pick_empalme_framing)

        def _wire_sup_diam_refresh(cmb):
            if cmb is None:
                return
            try:
                from System.Windows.Controls import SelectionChangedEventHandler

                def _on_sup_diam_changed(sender, args):
                    try:
                        self._refresh_empalmes_panel_from_selection()
                    except Exception:
                        pass

                cmb.SelectionChanged += SelectionChangedEventHandler(
                    _on_sup_diam_changed
                )
            except Exception:
                pass

        for _nm in (
            "CmbSupDiam",
            "CmbSup2Diam",
            "CmbSup3Diam",
            "CmbSupleDiam",
            "CmbSuple2Diam",
            "CmbInfDiam",
            "CmbInf2Diam",
            "CmbInf3Diam",
            "CmbSupleInfDiam",
            "CmbLatDiam",
            "CmbEstriboExtDiam",
            "CmbEstriboCentDiam",
        ):
            _wire_sup_diam_refresh(self._win.FindName(_nm))

        tb_ns = self._win.FindName("TxtNumCapasSuperiores")
        if tb_ns is not None:
            try:

                def _lf_cap_sup(s, e):
                    _normalize_capas_textbox(tb_ns)
                    try:
                        self._sync_capas_layer_panels_visibility()
                    except Exception:
                        pass

                def _tc_cap_sup(s, e):
                    try:
                        self._sync_capas_layer_panels_visibility()
                    except Exception:
                        pass

                tb_ns.LostFocus += RoutedEventHandler(_lf_cap_sup)
                tb_ns.TextChanged += RoutedEventHandler(_tc_cap_sup)
            except Exception:
                pass
        bu_ns = self._win.FindName("BtnNumCapasSupUp")
        bd_ns = self._win.FindName("BtnNumCapasSupDown")
        if bu_ns is not None and tb_ns is not None:
            try:

                def _cap_sup_u(s, e):
                    _bump_capas_textbox(tb_ns, 1)
                    try:
                        self._sync_capas_layer_panels_visibility()
                    except Exception:
                        pass

                bu_ns.Click += RoutedEventHandler(_cap_sup_u)
            except Exception:
                pass
        if bd_ns is not None and tb_ns is not None:
            try:

                def _cap_sup_d(s, e):
                    _bump_capas_textbox(tb_ns, -1)
                    try:
                        self._sync_capas_layer_panels_visibility()
                    except Exception:
                        pass

                bd_ns.Click += RoutedEventHandler(_cap_sup_d)
            except Exception:
                pass

        tb_ni = self._win.FindName("TxtNumCapasInferiores")
        if tb_ni is not None:
            try:

                def _lf_cap_inf(s, e):
                    _normalize_capas_textbox(tb_ni)
                    try:
                        self._sync_capas_layer_panels_visibility_inferior()
                    except Exception:
                        pass

                def _tc_cap_inf(s, e):
                    try:
                        self._sync_capas_layer_panels_visibility_inferior()
                    except Exception:
                        pass

                tb_ni.LostFocus += RoutedEventHandler(_lf_cap_inf)
                tb_ni.TextChanged += RoutedEventHandler(_tc_cap_inf)
            except Exception:
                pass
        bu_ni = self._win.FindName("BtnNumCapasInfUp")
        bd_ni = self._win.FindName("BtnNumCapasInfDown")
        if bu_ni is not None and tb_ni is not None:
            try:

                def _cap_inf_u(s, e):
                    _bump_capas_textbox(tb_ni, 1)
                    try:
                        self._sync_capas_layer_panels_visibility_inferior()
                    except Exception:
                        pass

                bu_ni.Click += RoutedEventHandler(_cap_inf_u)
            except Exception:
                pass
        if bd_ni is not None and tb_ni is not None:
            try:

                def _cap_inf_d(s, e):
                    _bump_capas_textbox(tb_ni, -1)
                    try:
                        self._sync_capas_layer_panels_visibility_inferior()
                    except Exception:
                        pass

                bd_ni.Click += RoutedEventHandler(_cap_inf_d)
            except Exception:
                pass

        tb_ss = self._win.FindName("TxtNumCapasSupleSup")
        if tb_ss is not None:
            try:

                def _lf_suple_sup(s, e):
                    _normalize_capas_suple_textbox(tb_ss)
                    try:
                        self._sync_suple_capa_labels()
                    except Exception:
                        pass

                def _tc_suple_sup(s, e):
                    try:
                        self._sync_suple_capa_labels()
                    except Exception:
                        pass

                tb_ss.LostFocus += RoutedEventHandler(_lf_suple_sup)
                tb_ss.TextChanged += RoutedEventHandler(_tc_suple_sup)
            except Exception:
                pass
        bu_ss = self._win.FindName("BtnNumCapasSupleSupUp")
        bd_ss = self._win.FindName("BtnNumCapasSupleSupDown")
        if bu_ss is not None and tb_ss is not None:
            try:

                def _suple_sup_u(s, e):
                    _bump_capas_suple_textbox(tb_ss, 1)
                    try:
                        self._sync_suple_capa_labels()
                    except Exception:
                        pass

                bu_ss.Click += RoutedEventHandler(_suple_sup_u)
            except Exception:
                pass
        if bd_ss is not None and tb_ss is not None:
            try:

                def _suple_sup_d(s, e):
                    _bump_capas_suple_textbox(tb_ss, -1)
                    try:
                        self._sync_suple_capa_labels()
                    except Exception:
                        pass

                bd_ss.Click += RoutedEventHandler(_suple_sup_d)
            except Exception:
                pass

        for tb_nm, up_nm, dn_nm in _CANTIDAD_SPINNER_TRIPLES:
            tb = self._win.FindName(tb_nm)
            bu = self._win.FindName(up_nm)
            bd = self._win.FindName(dn_nm)
            if tb is not None:
                try:
                    from System.Windows.Input import TextCompositionEventHandler

                    tb.PreviewTextInput += TextCompositionEventHandler(
                        _cantidad_tb_preview_text_input
                    )
                except Exception:
                    pass
                try:
                    from System.Windows import DataObject
                    from System.Windows import DataObjectPastingEventHandler

                    DataObject.AddPastingHandler(
                        tb,
                        DataObjectPastingEventHandler(_cantidad_tb_pasting),
                    )
                except Exception:
                    pass
                try:

                    def _on_lost_focus(s, e, tbx=tb):
                        _normalize_cantidad_textbox(tbx)

                    tb.LostFocus += RoutedEventHandler(_on_lost_focus)
                except Exception:
                    pass
            if bu is not None and tb is not None:
                try:

                    def _on_up(s, e, tbx=tb):
                        _bump_cantidad_textbox(tbx, 1)

                    bu.Click += RoutedEventHandler(_on_up)
                except Exception:
                    pass
            if bd is not None and tb is not None:
                try:

                    def _on_dn(s, e, tbx=tb):
                        _bump_cantidad_textbox(tbx, -1)

                    bd.Click += RoutedEventHandler(_on_dn)
                except Exception:
                    pass

        for tb_nm, up_nm, dn_nm in _ESTRIBO_SEP_SPINNER_TRIPLES:
            tb = self._win.FindName(tb_nm)
            bu = self._win.FindName(up_nm)
            bd = self._win.FindName(dn_nm)
            if tb is not None:
                try:
                    from System.Windows.Input import TextCompositionEventHandler

                    tb.PreviewTextInput += TextCompositionEventHandler(
                        _cantidad_tb_preview_text_input
                    )
                except Exception:
                    pass
                try:
                    from System.Windows import DataObject
                    from System.Windows import DataObjectPastingEventHandler

                    DataObject.AddPastingHandler(
                        tb,
                        DataObjectPastingEventHandler(_cantidad_tb_pasting),
                    )
                except Exception:
                    pass
                try:

                    def _on_lf_est_sep(s, e, tbx=tb):
                        _normalize_estribo_separacion_textbox(tbx)

                    tb.LostFocus += RoutedEventHandler(_on_lf_est_sep)
                except Exception:
                    pass
            if bu is not None and tb is not None:
                try:

                    def _on_est_sep_up(s, e, tbx=tb):
                        _bump_estribo_separacion_textbox(tbx, 10)

                    bu.Click += RoutedEventHandler(_on_est_sep_up)
                except Exception:
                    pass
            if bd is not None and tb is not None:
                try:

                    def _on_est_sep_dn(s, e, tbx=tb):
                        _bump_estribo_separacion_textbox(tbx, -10)

                    bd.Click += RoutedEventHandler(_on_est_sep_dn)
                except Exception:
                    pass

        try:
            self._poblar_combos_tipo_estribo()
        except Exception:
            pass

        try:
            from System.Windows.Input import MouseButtonEventHandler

            title_bar = self._win.FindName("TitleBar")
            if title_bar is not None:
                title_bar.MouseLeftButtonDown += MouseButtonEventHandler(
                    lambda s, e: self._win.DragMove()
                )
            if btn_close is not None:
                btn_close.MouseLeftButtonDown += MouseButtonEventHandler(lambda s, e: setattr(e, "Handled", True))
        except Exception:
            pass

        win_self = self

        for chk_name, panel_name in (
            ("ChkSuperior", "PanelSuperior"),
            ("ChkInferior", "PanelInferior"),
        ):
            chk = self._win.FindName(chk_name)
            pnl = self._win.FindName(panel_name)
            if chk is None or pnl is None:
                continue

            def _make_toggle(panel, name_sup_inf):
                def _toggle(s, a):
                    try:
                        en = s.IsChecked == True
                        panel.IsEnabled = en
                        panel.Opacity = 1.0 if en else 0.35
                    except Exception:
                        pass
                    if name_sup_inf == "ChkSuperior":
                        try:
                            win_self._sync_suple_capa_labels()
                        except Exception:
                            pass
                        try:
                            win_self._refresh_empalmes_panel_from_selection()
                        except Exception:
                            pass
                    elif name_sup_inf == "ChkInferior":
                        try:
                            win_self._sync_suple_inf_edits_enabled()
                            win_self._refresh_empalmes_panel_from_selection()
                        except Exception:
                            pass

                return _toggle

            chk.Checked += RoutedEventHandler(_make_toggle(pnl, chk_name))
            chk.Unchecked += RoutedEventHandler(_make_toggle(pnl, chk_name))

        chk_suple = self._win.FindName("ChkColocarSuple")
        pnl_suple = self._win.FindName("PnlSupleEdits")
        chk_sup_init = self._win.FindName("ChkSuperior")
        pnl_sup_init = self._win.FindName("PanelSuperior")

        def _sync_suple_edits_enabled(s, a):
            try:
                en_sup = chk_sup_init is None or chk_sup_init.IsChecked == True
                en_sp = chk_suple is None or chk_suple.IsChecked == True
                en = bool(en_sup and en_sp)
                if pnl_suple is not None:
                    pnl_suple.IsEnabled = en
                pnl_nc = self._win.FindName("PnlSupleSupNumCapasEdits")
                if pnl_nc is not None:
                    pnl_nc.IsEnabled = en
            except Exception:
                pass

        if chk_sup_init is not None and pnl_sup_init is not None:
            try:
                pnl_sup_init.IsEnabled = chk_sup_init.IsChecked == True
            except Exception:
                pass

        if chk_suple is not None:
            chk_suple.Checked += RoutedEventHandler(_sync_suple_edits_enabled)
            chk_suple.Unchecked += RoutedEventHandler(_sync_suple_edits_enabled)
        if chk_sup_init is not None:
            chk_sup_init.Checked += RoutedEventHandler(_sync_suple_edits_enabled)
            chk_sup_init.Unchecked += RoutedEventHandler(_sync_suple_edits_enabled)
        try:
            _sync_suple_edits_enabled(None, None)
        except Exception:
            pass

        chk_suple_inf = self._win.FindName("ChkColocarSupleInf")
        if chk_suple_inf is not None:
            chk_suple_inf.Checked += RoutedEventHandler(
                self._sync_suple_inf_edits_enabled
            )
            chk_suple_inf.Unchecked += RoutedEventHandler(
                self._sync_suple_inf_edits_enabled
            )
        try:
            self._sync_suple_inf_edits_enabled()
        except Exception:
            pass

        chk_lat = self._win.FindName("ChkLaterales")
        pnl_lat = self._win.FindName("PnlLateralesParams")
        if chk_lat is not None and pnl_lat is not None:

            def _toggle_laterales(s, a):
                en = s.IsChecked == True
                try:
                    pnl_lat.IsEnabled = en
                    pnl_lat.Opacity = 1.0 if en else 0.35
                except Exception:
                    pass

            try:
                chk_lat.Checked += RoutedEventHandler(_toggle_laterales)
                chk_lat.Unchecked += RoutedEventHandler(_toggle_laterales)
                _toggle_laterales(chk_lat, None)
            except Exception:
                pass

        chk_est = self._win.FindName("ChkEstribos")
        pnl_est = self._win.FindName("PnlEstribosParams")
        if chk_est is not None and pnl_est is not None:

            def _toggle_estribos(s, a):
                en = s.IsChecked == True
                try:
                    pnl_est.IsEnabled = en
                    pnl_est.Opacity = 1.0 if en else 0.35
                except Exception:
                    pass

            try:
                chk_est.Checked += RoutedEventHandler(_toggle_estribos)
                chk_est.Unchecked += RoutedEventHandler(_toggle_estribos)
                _toggle_estribos(chk_est, None)
            except Exception:
                pass

        try:
            self._init_num_capas_superiores()
        except Exception:
            pass
        try:
            self._init_num_capas_suple_superior()
        except Exception:
            pass
        try:
            self._init_num_capas_inferiores()
        except Exception:
            pass

    def _wire_commands(self, RoutedEventHandler, ApplicationCommands, CommandBinding, KeyBinding, Key, ModifierKeys):
        try:
            self._win.CommandBindings.Add(
                CommandBinding(ApplicationCommands.Close, RoutedEventHandler(lambda s, e: self._close_with_fade()))
            )
            self._win.InputBindings.Add(
                KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None)
            )
        except Exception:
            pass

    def _close_with_fade(self):
        """Inverso de _show_with_fade: opacidad 1→0 y Top de posición actual → +slide (misma duración / EaseInOut)."""
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

            ease = QuadraticEase()
            ease.EasingMode = EasingMode.EaseInOut

            opacity_anim = DoubleAnimation()
            opacity_anim.From = float(self._win.Opacity)
            opacity_anim.To = 0.0
            opacity_anim.Duration = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_CLOSE_MS)))
            opacity_anim.EasingFunction = ease

            current_top = float(self._win.Top)
            target_top = current_top + float(_WINDOW_SLIDE_PX)
            top_anim = DoubleAnimation()
            top_anim.From = current_top
            top_anim.To = target_top
            top_anim.Duration = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_CLOSE_MS)))
            top_anim.EasingFunction = ease

            def _done(s, a):
                try:
                    self._win.Close()
                except Exception:
                    pass

            opacity_anim.Completed += _done
            self._win.BeginAnimation(self._win.OpacityProperty, opacity_anim)
            self._win.BeginAnimation(self._win.TopProperty, top_anim)
        except Exception:
            self._is_closing_with_fade = False
            try:
                self._win.Close()
            except Exception:
                pass

    def _show_with_fade(self):
        try:
            from System import TimeSpan
            from System.Windows import Duration
            from System.Windows.Media.Animation import DoubleAnimation, QuadraticEase, EasingMode

            self._win.BeginAnimation(self._win.OpacityProperty, None)
            self._win.BeginAnimation(self._win.TopProperty, None)
            self._win.Opacity = 0.0
            if not self._win.IsVisible:
                self._win.Show()
            try:
                self._win.UpdateLayout()
            except Exception:
                pass
            try:
                self._base_top = float(self._win.Top)
            except Exception:
                self._base_top = 0.0
            start_top = float(self._base_top) + float(_WINDOW_SLIDE_PX)
            try:
                self._win.Top = start_top
            except Exception:
                pass
            ease = QuadraticEase()
            ease.EasingMode = EasingMode.EaseInOut
            oa = DoubleAnimation()
            oa.From = 0.0
            oa.To = 1.0
            oa.Duration = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_OPEN_MS)))
            oa.EasingFunction = ease
            ta = DoubleAnimation()
            ta.From = start_top
            ta.To = float(self._base_top)
            ta.Duration = Duration(TimeSpan.FromMilliseconds(float(_WINDOW_OPEN_MS)))
            ta.EasingFunction = ease
            self._win.BeginAnimation(self._win.OpacityProperty, oa)
            self._win.BeginAnimation(self._win.TopProperty, ta)
            self._is_closing_with_fade = False
            self._win.Activate()
        except Exception:
            try:
                self._win.Opacity = 1.0
                if not self._win.IsVisible:
                    self._win.Show()
                self._win.Activate()
            except Exception:
                pass

    def _set_estado(self, msg):
        try:
            txt = self._win.FindName("TxtEstado")
            if txt is not None:
                txt.Text = msg or u""
        except Exception:
            pass

    @staticmethod
    def _linea_elemento_detalle(elem):
        try:
            eid = element_id_to_int(elem.Id)
        except Exception:
            return None
        try:
            nm = unicode(elem.Name) if getattr(elem, "Name", None) else u""
        except Exception:
            nm = u""
        try:
            cat = (
                unicode(elem.Category.Name)
                if elem.Category is not None
                else u""
            )
        except Exception:
            cat = u""
        nm = nm.strip() if nm else u""
        if nm and cat:
            return u"  • Id {0} — {1} ({2})".format(eid, nm, cat)
        if nm:
            return u"  • Id {0} — {1}".format(eid, nm)
        if cat:
            return u"  • Id {0} ({1})".format(eid, cat)
        return u"  • Id {0}".format(eid)

    def _leer_texto_combo(self, cmb):
        if cmb is None:
            return u""
        try:
            t = cmb.Text
            if t is not None:
                s = unicode(t).strip()
                if s:
                    return s
        except Exception:
            pass
        try:
            si = cmb.SelectedItem
            if si is not None:
                return unicode(si).strip()
        except Exception:
            pass
        return u""

    def _texto_task_dialog_colocar_armadura(self):
        """Texto único para el TaskDialog al pulsar «Colocar armadura»: selección, UI y nota."""
        bloques = []
        bloques.append(u"Resumen — Colocar armadura")
        bloques.append(u"")
        doc = self._document

        n_v = len(self._selected_element_ids or [])
        bloques.append(u"Elementos seleccionados (vigas / columnas / muros): {0}".format(n_v))
        if not self._selected_element_ids:
            bloques.append(u"  (ninguno)")
        else:
            max_v = 40
            for eid in self._selected_element_ids[:max_v]:
                try:
                    el = doc.GetElement(eid) if doc else None
                except Exception:
                    el = None
                if el is None:
                    bloques.append(
                        u"  • Id {0}".format(element_id_to_int(eid))
                    )
                else:
                    ln = self._linea_elemento_detalle(el)
                    if ln:
                        bloques.append(ln)
            resto = n_v - min(n_v, max_v)
            if resto > 0:
                bloques.append(
                    u"  … y {0} elemento(s) más (no mostrados).".format(resto)
                )
        bloques.append(u"")

        chk_sup = self._win.FindName("ChkSuperior")
        chk_inf = self._win.FindName("ChkInferior")
        sup_on = chk_sup is None or chk_sup.IsChecked == True
        inf_on = chk_inf is None or chk_inf.IsChecked == True
        chk_ml_td = self._win.FindName("ChkModelLines")
        ml_on = chk_ml_td is None or chk_ml_td.IsChecked == True
        bloques.append(
            u"Model lines (guías): {0}.".format(u"sí" if ml_on else u"no")
        )
        bloques.append(u"")

        bloques.append(u"— Armadura superior —")
        bloques.append(
            u"Grupo activo: {0}".format(u"sí" if sup_on else u"no")
        )
        if sup_on:
            n_emp = len(getattr(self, "_empalme_framing_ids", None) or [])
            bloques.append(
                u"Empalmes (troceo automático si trazo > 12 m): vigas elegidas: {0}.".format(n_emp)
            )
            tb_nc = self._win.FindName("TxtNumCapasSuperiores")
            try:
                cap_txt = unicode(tb_nc.Text).strip() if tb_nc is not None else u"1"
            except Exception:
                cap_txt = u"1"
            bloques.append(u"Nº de capas: {0}.".format(cap_txt or u"1"))
            bloques.append(
                u"  1.ª — Cant.: {0}, Ø: {1}".format(
                    self._leer_texto_combo(self._win.FindName("CmbSupCant")),
                    self._leer_texto_combo(self._win.FindName("CmbSupDiam")),
                )
            )
            if _parse_capas_superiores_ventana(self) >= 2:
                bloques.append(
                    u"  2.ª — Cant.: {0}, Ø: {1}".format(
                        self._leer_texto_combo(
                            self._win.FindName("CmbSup2Cant")
                        ),
                        self._leer_texto_combo(
                            self._win.FindName("CmbSup2Diam")
                        ),
                    )
                )
            if _parse_capas_superiores_ventana(self) >= 3:
                bloques.append(
                    u"  3.ª — Cant.: {0}, Ø: {1}".format(
                        self._leer_texto_combo(
                            self._win.FindName("CmbSup3Cant")
                        ),
                        self._leer_texto_combo(
                            self._win.FindName("CmbSup3Diam")
                        ),
                    )
                )
            chk_sp = self._win.FindName("ChkColocarSuple")
            sp_on = chk_sp is None or chk_sp.IsChecked == True
            bloques.append(
                u"Suple bajo última capa: {0}.".format(u"sí" if sp_on else u"no")
            )
            if sp_on:
                tb_ns = self._win.FindName("TxtNumCapasSupleSup")
                try:
                    ns_txt = (
                        unicode(tb_ns.Text).strip()
                        if tb_ns is not None
                        else u"1"
                    )
                except Exception:
                    ns_txt = u"1"
                bloques.append(
                    u"Suple superior — Nº de capas (suple): {0}.".format(
                        ns_txt or u"1"
                    )
                )
                bloques.append(
                    u"  1.ª fila suple — Cant.: {0}, Ø: {1}".format(
                        self._leer_texto_combo(
                            self._win.FindName("CmbSupleCant")
                        ),
                        self._leer_texto_combo(
                            self._win.FindName("CmbSupleDiam")
                        ),
                    )
                )
                if _parse_capas_suple_superior_ventana(self) >= 2:
                    bloques.append(
                        u"  2.ª fila suple — Cant.: {0}, Ø: {1}".format(
                            self._leer_texto_combo(
                                self._win.FindName("CmbSuple2Cant")
                            ),
                            self._leer_texto_combo(
                                self._win.FindName("CmbSuple2Diam")
                            ),
                        )
                    )
        bloques.append(u"")

        bloques.append(u"— Armadura inferior —")
        bloques.append(
            u"Grupo activo: {0}".format(u"sí" if inf_on else u"no")
        )
        if inf_on:
            tb_ni = self._win.FindName("TxtNumCapasInferiores")
            try:
                cap_i_txt = (
                    unicode(tb_ni.Text).strip() if tb_ni is not None else u"1"
                )
            except Exception:
                cap_i_txt = u"1"
            bloques.append(u"Nº de capas: {0}.".format(cap_i_txt or u"1"))
            bloques.append(
                u"  1.ª — Cant.: {0}, Ø: {1}".format(
                    self._leer_texto_combo(self._win.FindName("CmbInfCant")),
                    self._leer_texto_combo(self._win.FindName("CmbInfDiam")),
                )
            )
            if _parse_capas_inferiores_ventana(self) >= 2:
                bloques.append(
                    u"  2.ª — Cant.: {0}, Ø: {1}".format(
                        self._leer_texto_combo(
                            self._win.FindName("CmbInf2Cant")
                        ),
                        self._leer_texto_combo(
                            self._win.FindName("CmbInf2Diam")
                        ),
                    )
                )
            if _parse_capas_inferiores_ventana(self) >= 3:
                bloques.append(
                    u"  3.ª — Cant.: {0}, Ø: {1}".format(
                        self._leer_texto_combo(
                            self._win.FindName("CmbInf3Cant")
                        ),
                        self._leer_texto_combo(
                            self._win.FindName("CmbInf3Diam")
                        ),
                    )
                )
            chk_spi = self._win.FindName("ChkColocarSupleInf")
            sp_i_on = chk_spi is None or chk_spi.IsChecked == True
            bloques.append(
                u"Suple sobre última capa: {0}.".format(
                    u"sí" if sp_i_on else u"no"
                )
            )
            if sp_i_on:
                bloques.append(
                    u"  Cantidad: {0}, Diámetro: {1}".format(
                        self._leer_texto_combo(
                            self._win.FindName("CmbSupleInfCant")
                        ),
                        self._leer_texto_combo(
                            self._win.FindName("CmbSupleInfDiam")
                        ),
                    )
                )
        bloques.append(u"")
        chk_est = self._win.FindName("ChkEstribos")
        est_on = chk_est is not None and chk_est.IsChecked == True
        bloques.append(u"— Estribos —")
        bloques.append(u"Activo: {0}.".format(u"sí" if est_on else u"no"))
        if est_on:
            bloques.append(
                u"Tipo: {0}.".format(
                    self._leer_texto_combo(self._win.FindName("CmbEstriboTipo"))
                )
            )
            bloques.append(
                u"Extremos — Ø: {0}, sep.: {1} mm.".format(
                    self._leer_texto_combo(self._win.FindName("CmbEstriboExtDiam")),
                    self._leer_texto_combo(self._win.FindName("TxtEstriboExtSep")),
                )
            )
            bloques.append(
                u"Centrales — Ø: {0}, sep.: {1} mm.".format(
                    self._leer_texto_combo(self._win.FindName("CmbEstriboCentDiam")),
                    self._leer_texto_combo(self._win.FindName("TxtEstriboCentSep")),
                )
            )
        bloques.append(u"")
        chk_lat = self._win.FindName("ChkLaterales")
        lat_on = chk_lat is not None and chk_lat.IsChecked == True
        bloques.append(u"— Barras laterales —")
        bloques.append(u"Activo: {0}.".format(u"sí" if lat_on else u"no"))
        if lat_on:
            bloques.append(
                u"Cant.: {0}, Ø: {1}".format(
                    self._leer_texto_combo(self._win.FindName("CmbLatCant")),
                    self._leer_texto_combo(self._win.FindName("CmbLatDiam")),
                )
            )
        return u"\n".join(bloques)

    def _refresh_selection_text(self):
        pass

    def _refresh_laterales_cantidad_desde_seleccion(self):
        """
        Actualiza la cantidad inicial de barras laterales según la mayor altura (*h*)
        entre las vigas Structural Framing de la selección.
        """
        doc = getattr(self, "_document", None)
        tb = self._win.FindName("CmbLatCant")
        if tb is None or doc is None:
            return
        ids = getattr(self, "_selected_element_ids", None) or []
        if not ids:
            return
        from geometria_viga_cara_superior_detalle import (
            filtrar_solo_structural_framing,
        )

        framing = filtrar_solo_structural_framing(doc, ids)
        if not framing:
            return
        h_max = None
        for fm in framing:
            h = _altura_viga_structural_mm(fm, doc)
            if h is not None:
                h_max = h if h_max is None else max(h_max, h)
        if h_max is None:
            return
        n = _cantidad_laterales_inicial_desde_altura_mm(h_max)
        try:
            tb.Text = unicode(n)
        except Exception:
            pass

    def _sync_suple_capa_labels(self):
        try:
            from System.Windows import Visibility

            n_main = _parse_capas_superiores_ventana(self)
            n_si = _parse_capas_suple_superior_ventana(self)
            prev_si = getattr(self, "_last_suple_sup_n_sincronizado", None)
            tb = self._win.FindName("TxtSupleSupCapaLbl")
            if tb is not None:
                tb.Text = _suple_sup_layer_label(n_main, 0)
            tb2 = self._win.FindName("TxtSupleSup2CapaLbl")
            if tb2 is not None:
                tb2.Text = _suple_sup_layer_label(n_main, 1)
            p2 = self._win.FindName("PnlSupleCapa2")
            if p2 is not None:
                activa2 = n_si >= 2
                p2.Visibility = (
                    Visibility.Visible if activa2 else Visibility.Collapsed
                )
                try:
                    p2.IsEnabled = activa2
                except Exception:
                    pass
                for nm in ("CmbSuple2Cant", "CmbSuple2Diam"):
                    c = self._win.FindName(nm)
                    if c is not None:
                        try:
                            c.IsEnabled = activa2
                        except Exception:
                            pass
            if n_si >= 2 and (prev_si is None or prev_si < 2):
                self._refrescar_seleccion_combos_suple2()
            self._last_suple_sup_n_sincronizado = n_si
            tb_i = self._win.FindName("TxtSupleInfCapaLbl")
            if tb_i is not None:
                tb_i.Text = _suple_capa_label_text(
                    _parse_capas_inferiores_ventana(self)
                )
        except Exception:
            pass

    def _fill_cantidad_combos(self):
        for tb_name, _, _ in _CANTIDAD_SPINNER_TRIPLES:
            tb = self._win.FindName(tb_name)
            if tb is None:
                continue
            try:
                tb.Text = _CANTIDAD_BARRAS_DEFAULT_TXT
            except Exception:
                pass

    def _init_num_capas_superiores(self):
        tb = self._win.FindName("TxtNumCapasSuperiores")
        if tb is not None:
            try:
                if not unicode(tb.Text).strip():
                    tb.Text = u"1"
            except Exception:
                try:
                    tb.Text = u"1"
                except Exception:
                    pass
        self._sync_capas_layer_panels_visibility()

    def _init_num_capas_suple_superior(self):
        tb = self._win.FindName("TxtNumCapasSupleSup")
        if tb is not None:
            try:
                if not unicode(tb.Text).strip():
                    tb.Text = _CAPAS_SUPLE_DEFAULT_TXT
            except Exception:
                try:
                    tb.Text = _CAPAS_SUPLE_DEFAULT_TXT
                except Exception:
                    pass
        try:
            self._sync_suple_capa_labels()
        except Exception:
            pass

    def _sync_capas_layer_panels_visibility(self):
        try:
            from System.Windows import Visibility

            n = _parse_capas_superiores_ventana(self)
            prev = getattr(self, "_last_capas_n_sincronizado", None)
            rows = (
                ("PnlCapa2", "CmbSup2Cant", "CmbSup2Diam", 2),
                ("PnlCapa3", "CmbSup3Cant", "CmbSup3Diam", 3),
            )
            for pname, c_cant, c_diam, idx in rows:
                pnl = self._win.FindName(pname)
                activa = n >= idx
                if pnl is not None:
                    try:
                        pnl.Visibility = (
                            Visibility.Visible if activa else Visibility.Collapsed
                        )
                    except Exception:
                        pass
                    try:
                        pnl.IsEnabled = activa
                    except Exception:
                        pass
                    if activa:
                        try:
                            pnl.UpdateLayout()
                        except Exception:
                            pass
                for cname in (c_cant, c_diam):
                    cmb = self._win.FindName(cname)
                    if cmb is not None:
                        try:
                            cmb.IsEnabled = activa
                        except Exception:
                            pass
            self._last_capas_n_sincronizado = n
            if n >= 2 and (prev is None or n > prev):
                self._refrescar_seleccion_combos_capas_visibles(n)
            try:
                self._sync_suple_capa_labels()
            except Exception:
                pass
        except Exception:
            pass

    def _refrescar_seleccion_combos_suple2(self):
        c0 = self._win.FindName("CmbSupleCant")
        d0 = self._win.FindName("CmbSupleDiam")
        ci = self._win.FindName("CmbSuple2Cant")
        di = self._win.FindName("CmbSuple2Diam")
        if ci is not None and c0 is not None:
            try:
                s0 = unicode(c0.Text).strip()
                if s0:
                    ci.Text = s0
            except Exception:
                pass
        if di is not None and d0 is not None:
            try:
                nid = int(di.Items.Count)
                if nid > 0 and di.SelectedIndex < 0:
                    s0 = int(d0.SelectedIndex)
                    if 0 <= s0 < nid:
                        di.SelectedIndex = s0
            except Exception:
                pass

    def _refrescar_seleccion_combos_capas_visibles(self, n_capas):
        """
        Si una capa recién activa no tiene ítem seleccionado en diámetro/cantidad,
        copia desde la primera capa para evitar leer valores inconsistentes.
        """
        n = max(1, min(3, int(n_capas or 1)))
        pairs = (
            ("CmbSupCant", "CmbSupDiam"),
            ("CmbSup2Cant", "CmbSup2Diam"),
            ("CmbSup3Cant", "CmbSup3Diam"),
        )
        c0, d0 = self._win.FindName("CmbSupCant"), self._win.FindName("CmbSupDiam")
        for i in range(1, n):
            ci = self._win.FindName(pairs[i][0])
            di = self._win.FindName(pairs[i][1])
            if ci is not None and c0 is not None:
                try:
                    s0 = unicode(c0.Text).strip()
                    if s0:
                        ci.Text = s0
                except Exception:
                    pass
            if di is not None and d0 is not None:
                try:
                    nid = int(di.Items.Count)
                    if nid > 0 and di.SelectedIndex < 0:
                        s0 = int(d0.SelectedIndex)
                        if 0 <= s0 < nid:
                            di.SelectedIndex = s0
                except Exception:
                    pass

    def _preparar_lectura_capas_superiores(self):
        """
        Antes de colocar: visibilidad, habilitación y layout alineados con ``TxtNumCapasSuperiores``;
        coherencia de combos de cada fila dinámica para leer parámetros correctos.
        """
        self._sync_capas_layer_panels_visibility()
        n = _parse_capas_superiores_ventana(self)
        if n >= 2:
            self._refrescar_seleccion_combos_capas_visibles(n)

    def _set_num_capas_superiores(self, n):
        n = max(1, min(3, int(n)))
        tb = self._win.FindName("TxtNumCapasSuperiores")
        if tb is not None:
            try:
                tb.Text = unicode(n)
            except Exception:
                pass
        self._sync_capas_layer_panels_visibility()

    def _sync_suple_inf_edits_enabled(self, sender=None, args=None):
        pnl = self._win.FindName("PnlSupleInfEdits")
        if pnl is None:
            return
        try:
            chk_inf = self._win.FindName("ChkInferior")
            chk_sp = self._win.FindName("ChkColocarSupleInf")
            from_enabled = chk_inf is None or chk_inf.IsChecked == True
            sup_on = chk_sp is None or chk_sp.IsChecked == True
            pnl.IsEnabled = bool(from_enabled and sup_on)
        except Exception:
            pass

    def _sync_capas_layer_panels_visibility_inferior(self):
        try:
            from System.Windows import Visibility

            n = _parse_capas_inferiores_ventana(self)
            prev = getattr(self, "_last_capas_inf_n_sincronizado", None)
            rows = (
                ("PnlInfCapa2", "CmbInf2Cant", "CmbInf2Diam", 2),
                ("PnlInfCapa3", "CmbInf3Cant", "CmbInf3Diam", 3),
            )
            for pname, c_cant, c_diam, idx in rows:
                pnl = self._win.FindName(pname)
                activa = n >= idx
                if pnl is not None:
                    try:
                        pnl.Visibility = (
                            Visibility.Visible if activa else Visibility.Collapsed
                        )
                    except Exception:
                        pass
                    try:
                        pnl.IsEnabled = activa
                    except Exception:
                        pass
                    if activa:
                        try:
                            pnl.UpdateLayout()
                        except Exception:
                            pass
                for cname in (c_cant, c_diam):
                    cmb = self._win.FindName(cname)
                    if cmb is not None:
                        try:
                            cmb.IsEnabled = activa
                        except Exception:
                            pass
            self._last_capas_inf_n_sincronizado = n
            if n >= 2 and (prev is None or n > prev):
                self._refrescar_seleccion_combos_capas_visibles_inferior(n)
            try:
                self._sync_suple_capa_labels()
            except Exception:
                pass
        except Exception:
            pass

    def _refrescar_seleccion_combos_capas_visibles_inferior(self, n_capas):
        n = max(1, min(3, int(n_capas or 1)))
        pairs = (
            ("CmbInfCant", "CmbInfDiam"),
            ("CmbInf2Cant", "CmbInf2Diam"),
            ("CmbInf3Cant", "CmbInf3Diam"),
        )
        c0, d0 = self._win.FindName("CmbInfCant"), self._win.FindName(
            "CmbInfDiam"
        )
        for i in range(1, n):
            ci = self._win.FindName(pairs[i][0])
            di = self._win.FindName(pairs[i][1])
            if ci is not None and c0 is not None:
                try:
                    s0 = unicode(c0.Text).strip()
                    if s0:
                        ci.Text = s0
                except Exception:
                    pass
            if di is not None and d0 is not None:
                try:
                    nid = int(di.Items.Count)
                    if nid > 0 and di.SelectedIndex < 0:
                        s0 = int(d0.SelectedIndex)
                        if 0 <= s0 < nid:
                            di.SelectedIndex = s0
                except Exception:
                    pass

    def _preparar_lectura_capas_inferiores(self):
        self._sync_capas_layer_panels_visibility_inferior()
        n = _parse_capas_inferiores_ventana(self)
        if n >= 2:
            self._refrescar_seleccion_combos_capas_visibles_inferior(n)

    def _set_num_capas_inferiores(self, n):
        n = max(1, min(3, int(n)))
        tb = self._win.FindName("TxtNumCapasInferiores")
        if tb is not None:
            try:
                tb.Text = unicode(n)
            except Exception:
                pass
        self._sync_capas_layer_panels_visibility_inferior()

    def _init_num_capas_inferiores(self):
        tb = self._win.FindName("TxtNumCapasInferiores")
        if tb is not None:
            try:
                if not unicode(tb.Text).strip():
                    tb.Text = u"1"
            except Exception:
                try:
                    tb.Text = u"1"
                except Exception:
                    pass
        self._sync_capas_layer_panels_visibility_inferior()

    def _poblar_combos_tipo_estribo(self):
        for nm in (u"CmbEstriboTipo",):
            cmb = self._win.FindName(nm)
            if cmb is None:
                continue
            try:
                cmb.Items.Clear()
                cmb.IsEditable = False
                for s in (u"Simple", u"Doble", u"Triple"):
                    cmb.Items.Add(s)
                try:
                    cmb.SelectedIndex = 0
                except Exception:
                    pass
            except Exception:
                pass

    def _cargar_combos_diametro(self):
        doc = self._document
        labels = _etiquetas_diametro_estandar()
        self._entries = [(None, lbl) for lbl in labels]
        if doc is None:
            return
        for name in (
            "CmbSupDiam",
            "CmbSup2Diam",
            "CmbSup3Diam",
            "CmbSupleDiam",
            "CmbSuple2Diam",
            "CmbInfDiam",
            "CmbInf2Diam",
            "CmbInf3Diam",
            "CmbSupleInfDiam",
            "CmbLatDiam",
            "CmbEstriboExtDiam",
            "CmbEstriboCentDiam",
        ):
            cmb = self._win.FindName(name)
            if cmb is None:
                continue
            cmb.Items.Clear()
            cmb.IsEditable = False
            for lbl in labels:
                cmb.Items.Add(lbl)
            try:
                cmb.SelectedIndex = 0
            except Exception:
                pass

    def _on_seleccionar(self, sender, args):
        try:
            self._win.Hide()
        except Exception:
            pass
        self._seleccion_event.Raise()

    def _set_empalmes_panel_visible(self, visible):
        pnl = self._win.FindName("PnlEmpalmesSup")
        if pnl is None:
            return
        try:
            from System.Windows import Visibility

            pnl.Visibility = Visibility.Visible if visible else Visibility.Collapsed
        except Exception:
            try:
                pnl.Visibility = u"Visible" if visible else u"Collapsed"
            except Exception:
                pass

    def _sync_empalme_pick_button_enabled(self):
        btn = self._win.FindName("BtnVigasEmpalme")
        pnl = self._win.FindName("PnlEmpalmesSup")
        if btn is None:
            return
        try:
            from System.Windows import Visibility

            vis = (
                pnl is not None
                and pnl.Visibility == Visibility.Visible
            )
            btn.IsEnabled = bool(vis)
        except Exception:
            try:
                btn.IsEnabled = pnl is not None
            except Exception:
                pass

    def _reset_empalmes_ui_state(self):
        self._empalme_framing_ids = []
        self._refresh_empalme_text()
        self._sync_empalme_pick_button_enabled()

    def _estimar_largo_max_trazo_mm_para_empalmes(self):
        """
        Máximo del largo físico estimado (mm, eje + ganchos) entre caras superior/inferior
        activas en la UI; ``None`` si no hay documento, selección, vigas u orientación.
        """
        chk_sup = self._win.FindName("ChkSuperior")
        chk_inf = self._win.FindName("ChkInferior")
        sup_on = chk_sup is None or chk_sup.IsChecked == True
        inf_on = chk_inf is None or chk_inf.IsChecked == True
        if not sup_on and not inf_on:
            return None
        doc = self._document
        if doc is not None and not (getattr(self, "_entries", None) or []):
            try:
                self._cargar_combos_diametro()
            except Exception:
                pass
        ids = getattr(self, "_selected_element_ids", None) or []
        if doc is None or not ids:
            return None
        from geometria_viga_cara_superior_detalle import (
            estimar_largo_mm_trazo_inferior,
            estimar_largo_mm_trazo_superior,
            filtrar_obstaculos_seleccion_no_framing,
            filtrar_solo_structural_framing,
        )

        framing = filtrar_solo_structural_framing(doc, ids)
        obst = filtrar_obstaculos_seleccion_no_framing(doc, ids)
        if not framing:
            return None
        L_mm = None
        if sup_on:
            try:
                bar_sup = _rebar_bar_type_superior_desde_ventana(self)
            except Exception:
                bar_sup = None
            try:
                Ls = estimar_largo_mm_trazo_superior(
                    doc, framing, obst, bar_sup
                )
                if Ls is not None:
                    L_mm = (
                        float(Ls)
                        if L_mm is None
                        else max(float(L_mm), float(Ls))
                    )
            except Exception:
                pass
        if inf_on:
            try:
                self._preparar_lectura_capas_inferiores()
            except Exception:
                pass
            try:
                nci = _parse_capas_inferiores_ventana(self)
                ti, _ = _inferior_rebar_types_y_cantidades_por_capa(self, nci)
                bar_inf = ti[0] if ti else None
            except Exception:
                bar_inf = None
            try:
                Li = estimar_largo_mm_trazo_inferior(
                    doc, framing, obst, bar_inf
                )
                if Li is not None:
                    L_mm = (
                        float(Li)
                        if L_mm is None
                        else max(float(L_mm), float(Li))
                    )
            except Exception:
                pass
        return L_mm

    def _traslape_empalme_obligatorio_sin_vigas_definidas(self):
        """True si el trazo supera 12 m y no hay ninguna viga Structural Framing de empalme."""
        from geometria_viga_cara_superior_detalle import (
            _UMBRAL_EMPALMES_COMP_EPS_MM,
            _UMBRAL_LARGO_EMPALMES_MM,
        )

        L_mm = self._estimar_largo_max_trazo_mm_para_empalmes()
        if L_mm is None:
            return False
        umbral = float(_UMBRAL_LARGO_EMPALMES_MM) - float(
            _UMBRAL_EMPALMES_COMP_EPS_MM
        )
        if float(L_mm) <= umbral:
            return False
        emp = getattr(self, "_empalme_framing_ids", None) or []
        return len(emp) == 0

    def _refresh_empalmes_panel_from_selection(self):
        """Muestra el bloque empalmes si el trazo superior o inferior estimado supera 12 m."""
        L_mm = self._estimar_largo_max_trazo_mm_para_empalmes()
        if L_mm is None:
            self._set_empalmes_panel_visible(False)
            self._reset_empalmes_ui_state()
            return
        from geometria_viga_cara_superior_detalle import (
            _UMBRAL_EMPALMES_COMP_EPS_MM,
            _UMBRAL_LARGO_EMPALMES_MM,
        )

        umbral = float(_UMBRAL_LARGO_EMPALMES_MM) - float(
            _UMBRAL_EMPALMES_COMP_EPS_MM
        )
        if float(L_mm) > umbral:
            self._set_empalmes_panel_visible(True)
            self._sync_empalme_pick_button_enabled()
        else:
            self._set_empalmes_panel_visible(False)
            self._reset_empalmes_ui_state()

    def _refresh_empalme_text(self):
        """La selección de vigas de empalme no muestra listado de Id en la UI."""
        pass

    def _on_pick_empalme_framing(self, sender, args):
        try:
            self._win.Hide()
        except Exception:
            pass
        self._empalme_pick_event.Raise()

    def _on_colocar(self, sender, args):
        if not self._selected_element_ids:
            _task_dialog_show(
                u"BIMTools — Enfierrado vigas",
                u"Seleccione al menos un elemento (viga, columna o muro) en el modelo.",
                self._win,
            )
            self._set_estado(u"Seleccione elementos antes de continuar.")
            return
        chk_inf = self._win.FindName("ChkInferior")
        chk_sup = self._win.FindName("ChkSuperior")
        if chk_inf and chk_sup:
            if chk_inf.IsChecked != True and chk_sup.IsChecked != True:
                _task_dialog_show(
                    u"BIMTools — Enfierrado vigas",
                    u"Active al menos un grupo de armadura (superior o inferior).",
                    self._win,
                )
                return
        if self._traslape_empalme_obligatorio_sin_vigas_definidas():
            _task_dialog_show(
                u"BIMTools — Enfierrado vigas — Traslape / empalme",
                u"El trazo estimado supera 12 m (eje y ganchos). Debe definir las vigas de "
                u"empalme antes de colocar la armadura.\n\n"
                u"1) Pulse «Seleccionar vigas para realizar empalmes» (bloque visible bajo la selección).\n"
                u"2) Elija al menos una viga Structural Framing: en ellas se sitúan los planos de troceo y traslape.\n"
                u"3) Pulse «Colocar armadura» de nuevo.\n\n"
                u"Sin esa definición no se aplicará el troceo por juntas de empalme.",
                self._win,
            )
            try:
                self._set_empalmes_panel_visible(True)
                self._sync_empalme_pick_button_enabled()
            except Exception:
                pass
            try:
                self._set_estado(
                    u"Traslape obligatorio: defina vigas de empalme (trazo > 12 m)."
                )
            except Exception:
                pass
            return
        self._colocar_event.Raise()
        self._set_estado(
            u"En cola: armadura longitudinal (cara superior y/o inferior)…"
        )

    def show(self):
        uidoc = self._revit.ActiveUIDocument
        if uidoc is None:
            _task_dialog_show(
                u"Enfierrado vigas",
                u"No hay documento activo.",
                self._win,
            )
            return
        self._document = uidoc.Document
        hwnd = None
        try:
            hwnd = revit_main_hwnd(self._revit.Application)
        except Exception:
            pass
        try:
            from System.Windows.Interop import WindowInteropHelper

            if hwnd:
                helper = WindowInteropHelper(self._win)
                helper.Owner = hwnd
        except Exception:
            pass
        position_wpf_window_top_left_at_active_view(self._win, uidoc, hwnd)
        self._fill_cantidad_combos()
        self._cargar_combos_diametro()
        try:
            self._init_num_capas_superiores()
        except Exception:
            pass
        try:
            self._init_num_capas_suple_superior()
        except Exception:
            pass
        try:
            self._init_num_capas_inferiores()
        except Exception:
            pass
        try:
            self._sync_suple_inf_edits_enabled()
        except Exception:
            pass
        try:
            self._refresh_empalmes_panel_from_selection()
        except Exception:
            pass
        self._set_estado(
            u"Pulse «Seleccionar en Modelo» para elegir vigas, columnas o muros."
        )
        self._show_with_fade()
        try:
            System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, self._win)
        except Exception:
            pass


def run_pyrevit(revit):
    if _scripts_dir not in sys.path:
        sys.path.insert(0, _scripts_dir)

    existing = _get_active_window()
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
            _clear_appdomain_window_key()
            existing = None
        if ok and existing is not None:
            _task_dialog_show(
                u"BIMTools — Enfierrado vigas",
                u"La herramienta ya está en ejecución.",
                existing,
            )
            return

    w = EnfierradoVigasWindow(revit)
    try:
        w.show()
    except Exception:
        _clear_appdomain_window_key()
        raise
