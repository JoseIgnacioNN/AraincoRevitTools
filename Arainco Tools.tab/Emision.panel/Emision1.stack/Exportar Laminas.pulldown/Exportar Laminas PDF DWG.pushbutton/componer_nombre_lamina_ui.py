# -*- coding: utf-8 -*-
"""
Diálogo modal: Nombre Personalizado — composición del nombre de archivo por parámetros de lámina.

La receta confirmada con «Aplicar nombres» se guarda en el documento (Extensible Storage en
información del proyecto) para reutilizarla al abrir el mismo modelo en otra sesión o
tras sincronizar (worksharing / ACC).
"""

import imp
import os
import sys

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import ElementId, ViewSheet  # noqa: E402
from System import EventHandler, String  # noqa: E402
from System.Data import DataColumn, DataTable  # noqa: E402
from System.Windows import (  # noqa: E402
    RoutedEventHandler,
    SizeChangedEventHandler,
    Visibility,
)
from System.Windows.Controls import (  # noqa: E402
    DataGridCellEditEndingEventArgs,
    DataGridEditingUnit,
    TextChangedEventHandler,
)
from System.Windows.Markup import XamlReader  # noqa: E402

_pb = os.path.dirname(os.path.abspath(__file__))
if _pb not in sys.path:
    sys.path.insert(0, _pb)
_d = _pb
for _ in range(24):
    _sp = os.path.join(_d, "scripts")
    if os.path.isfile(os.path.join(_sp, "bimtools_wpf_dark_theme.py")):
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
            "bimtools_paths__ComponerNombreLamina", _bimtools_paths_fp
        )
    else:
        import bimtools_paths as _bimtools_paths_std

        bimtools_paths = _bimtools_paths_std
    bimtools_paths.set_pushbutton_dir(_pb)
except Exception:
    bimtools_paths = None

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""


def _build_options_datatable(opts_list):
    import clr as _clr

    dt = DataTable()
    dt.Columns.Add(DataColumn(u"Key", _clr.GetClrType(String)))
    dt.Columns.Add(DataColumn(u"Label", _clr.GetClrType(String)))
    for o in opts_list or []:
        r = dt.NewRow()
        r[u"Key"] = o.get(u"Key", u"")
        r[u"Label"] = o.get(u"Label", u"")
        dt.Rows.Add(r)
    return dt


def _nullable_wpf_bool(nb):
    try:
        if nb is None:
            return False
        if hasattr(nb, u"HasValue"):
            return bool(nb.HasValue and nb.Value)
        return bool(nb)
    except Exception:
        return False


def _recipe_segments_from_table(recipe_dt):
    segs = []
    for i in range(recipe_dt.Rows.Count):
        row = recipe_dt.Rows[i]
        try:
            segs.append(
                {
                    u"Key": unicode(row[u"Key"]),
                    u"Prefix": unicode(row[u"Prefix"] or u""),
                    u"Suffix": unicode(row[u"Suffix"] or u""),
                    u"Separator": unicode(row[u"Separator"] or u""),
                }
            )
        except Exception:
            continue
    return segs


def _sample_sheet(doc, main_table):
    """Primera lámina con checkbox o, si no hay, la primera fila."""
    from Autodesk.Revit.DB import ElementId as _Eid

    n = main_table.Rows.Count
    for i in range(n):
        row = main_table.Rows[i]
        try:
            if not _nullable_wpf_bool(row[u"Sel"]):
                continue
        except Exception:
            continue
        try:
            sid = int(row[u"IdInt"])
        except Exception:
            continue
        try:
            el = doc.GetElement(_Eid(sid))
            if isinstance(el, ViewSheet):
                return el
        except Exception:
            pass
    for i in range(n):
        row = main_table.Rows[i]
        try:
            sid = int(row[u"IdInt"])
        except Exception:
            continue
        try:
            el = doc.GetElement(_Eid(sid))
            if isinstance(el, ViewSheet):
                return el
        except Exception:
            pass
    return None


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


XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="NbNamingWin"
    Title="BIMTools — Nombre Personalizado"
    Height="600" Width="980" MinHeight="520" MinWidth="800"
    Background="Transparent"
    AllowsTransparency="True"
    WindowStyle="None"
    ResizeMode="CanResize"
    WindowStartupLocation="CenterOwner"
    UseLayoutRounding="True"
    FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
""" + BIMTOOLS_DARK_STYLES_XML + u"""
    <Style x:Key="StepBadge" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#7ED8ED"/>
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
      <Setter Property="Foreground" Value="#C8E4EF"/>
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
    <Style x:Key="BtnTransfer" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Padding" Value="0,0"/>
      <Setter Property="Width" Value="44"/>
      <Setter Property="Height" Value="44"/>
      <Setter Property="FontWeight" Value="Bold"/>
    </Style>
    <Style x:Key="BtnTransferOut" TargetType="Button" BasedOn="{StaticResource BtnSelectOutline}">
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Padding" Value="0,0"/>
      <Setter Property="Width" Value="44"/>
      <Setter Property="Height" Value="44"/>
      <Setter Property="FontWeight" Value="Bold"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#050E18"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#284760"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="FontSize" Value="11"/>
    </Style>
    <Style x:Key="DgEditLeft" TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
      <Setter Property="TextAlignment" Value="Left"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="#F2F8FC"/>
      <Setter Property="CaretBrush" Value="#5BC0DE"/>
    </Style>
    <Style x:Key="DgTbLeft" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#F2F8FC"/>
      <Setter Property="TextAlignment" Value="Left"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="NbDgHeaderDark" TargetType="DataGridColumnHeader">
      <Setter Property="OverridesDefaultStyle" Value="True"/>
      <Setter Property="Background" Value="#122A3A"/>
      <Setter Property="Foreground" Value="#F2F8FC"/>
      <Setter Property="BorderBrush" Value="#1E3F55"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
      <Setter Property="Padding" Value="12,10"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridColumnHeader">
            <Grid>
              <Border x:Name="HeaderBorder" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True"
                      Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}"
                      TextElement.Foreground="{TemplateBinding Foreground}">
                <ContentPresenter RecognizesAccessKey="True" SnapsToDevicePixels="True"
                                  HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                  VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
              </Border>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="HeaderBorder" Property="Background" Value="#1B405C"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="HeaderBorder" Property="Background" Value="#0F2840"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="DgHdrCenter" TargetType="DataGridColumnHeader" BasedOn="{StaticResource NbDgHeaderDark}">
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="DgHdrLeft" TargetType="DataGridColumnHeader" BasedOn="{StaticResource NbDgHeaderDark}">
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
    </Style>
    <Style x:Key="OpcionesListBoxStyle" TargetType="ListBox">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="#F2F8FC"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="4,2"/>
      <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Visible"/>
      <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
      <Setter Property="VirtualizingStackPanel.IsVirtualizing" Value="False"/>
      <Setter Property="ItemContainerStyle">
        <Setter.Value>
          <Style TargetType="ListBoxItem">
            <Setter Property="Foreground" Value="#F2F8FC"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Margin" Value="0,2"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Template">
              <Setter.Value>
                <ControlTemplate TargetType="ListBoxItem">
                  <Border x:Name="Bd" Background="Transparent" BorderBrush="Transparent" BorderThickness="1"
                          CornerRadius="4" Padding="8,6" SnapsToDevicePixels="True"
                          TextElement.Foreground="{TemplateBinding Foreground}">
                    <ContentPresenter/>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                      <Setter TargetName="Bd" Property="Background" Value="#152A40"/>
                    </Trigger>
                    <Trigger Property="IsSelected" Value="True">
                      <Setter TargetName="Bd" Property="BorderBrush" Value="#5BC0DE"/>
                      <Setter TargetName="Bd" Property="Background" Value="#132A40"/>
                    </Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </Setter.Value>
            </Setter>
          </Style>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="DataGrid" BasedOn="{StaticResource {x:Type DataGrid}}">
      <Setter Property="Background" Value="#040A12"/>
      <Setter Property="Foreground" Value="#F2F8FC"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="RowBackground" Value="#0B1726"/>
      <Setter Property="AlternatingRowBackground" Value="#071420"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#152A3D"/>
      <Setter Property="VerticalGridLinesBrush" Value="#152A3D"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="RowHeight" Value="34"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource NbDgHeaderDark}"/>
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
      <Setter Property="Foreground" Value="#F2F8FC"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="Transparent"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>
  <Border x:Name="NbRootChrome" CornerRadius="8" Background="#0E1B32" Padding="14"
          BorderBrush="#5BC0DE" BorderThickness="1" ClipToBounds="True">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="16" ShadowDepth="0" Opacity="0.35"/>
    </Border.Effect>
    <Grid>
    <Grid.RowDefinitions>
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
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <Image x:Name="ImgLogo" Width="40" Height="40"
                 Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center" RenderOptions.BitmapScalingMode="HighQuality"/>
          <TextBlock Text="Nombre Personalizado" FontSize="14" FontWeight="Bold" Foreground="#FFFFFF"
                     VerticalAlignment="Center"/>
        </StackPanel>
        <Button x:Name="BtnClose" Grid.Column="2" HorizontalAlignment="Right" VerticalAlignment="Top"
                Style="{StaticResource BtnCloseX_MinimalNoBg}" Padding="6" Margin="0,-4,-4,0"/>
      </Grid>
    </Border>
    <Grid Grid.Row="1" VerticalAlignment="Stretch">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="290"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Border Grid.Column="0" Style="{StaticResource PanelInset}" Margin="0,0,12,12" VerticalAlignment="Stretch">
        <Grid VerticalAlignment="Stretch">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <TextBlock Grid.Row="0" Text="Parámetros disponibles" Style="{StaticResource StepBadge}"/>
          <TextBox x:Name="TxtFiltroOpciones" Grid.Row="1" Margin="0,0,0,8" ToolTip="Filtrar por nombre de parámetro"/>
          <Border x:Name="BdOpcionesLista" Grid.Row="2" Background="#040A12" BorderBrush="#1E3F55" BorderThickness="1" CornerRadius="8"
                  MinHeight="200" VerticalAlignment="Stretch" SnapsToDevicePixels="True">
            <ListBox x:Name="LstOpciones" Style="{StaticResource OpcionesListBoxStyle}" DisplayMemberPath="Label"/>
          </Border>
        </Grid>
      </Border>
      <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="0,0,12,0" MinWidth="48">
        <Button x:Name="BtnAnadir" Content="&#x2192;" ToolTip="Añadir a la secuencia" Style="{StaticResource BtnTransfer}" Margin="0,0,0,10"/>
        <Button x:Name="BtnQuitar" Content="&#x2190;" ToolTip="Quitar de la secuencia" Style="{StaticResource BtnTransferOut}"/>
      </StackPanel>
      <Border Grid.Column="2" Style="{StaticResource PanelInset}" VerticalAlignment="Stretch">
        <Grid VerticalAlignment="Stretch">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock Grid.Row="0" Text="Secuencia del nombre" Style="{StaticResource StepBadge}"/>
          <DataGrid x:Name="GridReceta" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" CanUserSortColumns="False"
                    CanUserResizeColumns="False" SelectionMode="Single" MinHeight="140" Margin="0,0,0,10" GridLinesVisibility="Horizontal"
                    RowHeaderWidth="0" VerticalContentAlignment="Center">
            <DataGrid.CellStyle>
              <Style TargetType="DataGridCell" BasedOn="{StaticResource GridCellPadding}"/>
            </DataGrid.CellStyle>
            <DataGrid.Columns>
              <DataGridTextColumn Header="Parámetro" Binding="{Binding Origen}" Width="*" MinWidth="120" IsReadOnly="True">
                <DataGridTextColumn.HeaderStyle>
                  <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrLeft}"/>
                </DataGridTextColumn.HeaderStyle>
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbLeft}"/>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="Prefijo" Binding="{Binding Prefix, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="72">
                <DataGridTextColumn.HeaderStyle>
                  <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrCenter}"/>
                </DataGridTextColumn.HeaderStyle>
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbLeft}"/>
                </DataGridTextColumn.ElementStyle>
                <DataGridTextColumn.EditingElementStyle>
                  <Style TargetType="TextBox" BasedOn="{StaticResource DgEditLeft}"/>
                </DataGridTextColumn.EditingElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="Sufijo" Binding="{Binding Suffix, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="72">
                <DataGridTextColumn.HeaderStyle>
                  <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrCenter}"/>
                </DataGridTextColumn.HeaderStyle>
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbLeft}"/>
                </DataGridTextColumn.ElementStyle>
                <DataGridTextColumn.EditingElementStyle>
                  <Style TargetType="TextBox" BasedOn="{StaticResource DgEditLeft}"/>
                </DataGridTextColumn.EditingElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="Sep." Binding="{Binding Separator, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="56">
                <DataGridTextColumn.HeaderStyle>
                  <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrCenter}"/>
                </DataGridTextColumn.HeaderStyle>
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock" BasedOn="{StaticResource DgTbLeft}"/>
                </DataGridTextColumn.ElementStyle>
                <DataGridTextColumn.EditingElementStyle>
                  <Style TargetType="TextBox" BasedOn="{StaticResource DgEditLeft}"/>
                </DataGridTextColumn.EditingElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="" Binding="{Binding Key}" Width="0" MinWidth="0" MaxWidth="0" IsReadOnly="True">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Setter Property="Visibility" Value="Collapsed"/>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
            </DataGrid.Columns>
          </DataGrid>
          <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,12">
            <Button x:Name="BtnSubir" Content="Subir" Style="{StaticResource BtnGhost}" Margin="0,0,8,0" MinWidth="72"/>
            <Button x:Name="BtnBajar" Content="Bajar" Style="{StaticResource BtnGhost}" Margin="0,0,8,0" MinWidth="72"/>
            <Button x:Name="BtnLimpiar" Content="Limpiar secuencia" Style="{StaticResource BtnGhost}" MinWidth="132"/>
          </StackPanel>
          <TextBlock Grid.Row="3" Text="Vista previa" Style="{StaticResource StepBadge}"/>
          <Border Grid.Row="4" Background="#040A12" BorderBrush="#1E3F55" BorderThickness="1" CornerRadius="8"
                  Padding="14,12" MinHeight="88" VerticalAlignment="Stretch">
            <TextBlock x:Name="TxtVistaPrevia" Foreground="#F2F8FC" FontSize="13" FontWeight="SemiBold"
                       TextWrapping="Wrap" VerticalAlignment="Center"/>
          </Border>
        </Grid>
      </Border>
    </Grid>
    <Grid Grid.Row="2" Margin="0,16,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBlock Grid.Column="0" Foreground="#92B4C9" FontSize="11" VerticalAlignment="Center" TextWrapping="Wrap"
                 Text="Aplicar reemplaza la columna «Nombre de archivo» en todas las filas del diálogo principal."/>
      <Button x:Name="BtnOk" Grid.Column="1" Content="Aplicar nombres" Style="{StaticResource BtnPrimary}" MinWidth="148" Margin="0,0,10,0"/>
      <Button x:Name="BtnCancel" Grid.Column="2" Content="Cancelar" Style="{StaticResource BtnSelectOutline}" MinWidth="100"/>
    </Grid>
    </Grid>
  </Border>
</Window>
"""


def _hydrate_recipe_table(recipe_dt, doc, opts_list):
    """
    Rellena ``recipe_dt`` con la receta persistida en el documento (si hay y es válida).
    """
    try:
        import exportar_laminas_pdf_dwg as _ex_nm  # noqa: WPS433
    except Exception:
        return
    try:
        segs = _ex_nm.get_persisted_naming_recipe_segments(doc)
    except Exception:
        segs = []
    if not segs:
        return
    label_by = {}
    for o in opts_list or []:
        try:
            k = o.get(u"Key", u"")
            if k:
                label_by[unicode(k)] = unicode(o.get(u"Label", u""))
        except Exception:
            continue
    for seg in segs:
        try:
            key = unicode(seg.get(u"Key", u""))
        except Exception:
            key = u""
        if not key:
            continue
        row = recipe_dt.NewRow()
        row[u"Key"] = key
        row[u"Origen"] = label_by.get(key, key)
        try:
            row[u"Prefix"] = unicode(seg.get(u"Prefix") or u"")
            row[u"Suffix"] = unicode(seg.get(u"Suffix") or u"")
            row[u"Separator"] = unicode(seg.get(u"Separator") or u"")
        except Exception:
            row[u"Prefix"] = u""
            row[u"Suffix"] = u""
            row[u"Separator"] = u"_"
        try:
            recipe_dt.Rows.Add(row)
        except Exception:
            pass


def _swap_recipe_rows(dt, i, j):
    if i == j or i < 0 or j < 0 or i >= dt.Rows.Count or j >= dt.Rows.Count:
        return
    ri, rj = dt.Rows[i], dt.Rows[j]
    for c in dt.Columns:
        n = c.ColumnName
        tmp = ri[n]
        ri[n] = rj[n]
        rj[n] = tmp


class ComponerNombreLaminaDialog(object):
    def __init__(self, owner_wpf, doc, main_datatable, list_options_fn, evaluate_fn):
        self._doc = doc
        self._main = main_datatable
        self._evaluate = evaluate_fn
        self._opts_full = list_options_fn(doc)
        self._opts_dt = _build_options_datatable(self._opts_full)
        self._recipe = DataTable()
        import clr as _clr

        self._recipe.Columns.Add(DataColumn(u"Key", _clr.GetClrType(String)))
        self._recipe.Columns.Add(DataColumn(u"Origen", _clr.GetClrType(String)))
        self._recipe.Columns.Add(DataColumn(u"Prefix", _clr.GetClrType(String)))
        self._recipe.Columns.Add(DataColumn(u"Suffix", _clr.GetClrType(String)))
        self._recipe.Columns.Add(DataColumn(u"Separator", _clr.GetClrType(String)))

        _hydrate_recipe_table(self._recipe, self._doc, self._opts_full)

        self._win = XamlReader.Parse(XAML)
        _load_bimtools_logo_into_window(self._win)
        self._lst_opciones = self._win.FindName(u"LstOpciones")
        self._bd_opciones = self._win.FindName(u"BdOpcionesLista")
        self._txt_filtro = self._win.FindName(u"TxtFiltroOpciones")
        self._grid = self._win.FindName(u"GridReceta")
        self._txt_prev = self._win.FindName(u"TxtVistaPrevia")

        self._lst_opciones.ItemsSource = self._opts_dt.DefaultView
        self._grid.ItemsSource = self._recipe.DefaultView

        self._txt_filtro.TextChanged += TextChangedEventHandler(self._on_filtro_opciones)
        self._win.FindName(u"BtnAnadir").Click += RoutedEventHandler(self._on_anadir)
        self._win.FindName(u"BtnQuitar").Click += RoutedEventHandler(self._on_quitar)
        self._win.FindName(u"BtnSubir").Click += RoutedEventHandler(self._on_subir)
        self._win.FindName(u"BtnBajar").Click += RoutedEventHandler(self._on_bajar)
        self._win.FindName(u"BtnLimpiar").Click += RoutedEventHandler(self._on_limpiar)
        self._win.FindName(u"BtnOk").Click += RoutedEventHandler(self._on_ok)
        self._win.FindName(u"BtnCancel").Click += RoutedEventHandler(self._on_cancel)
        try:
            from System.Windows.Input import (
                ApplicationCommands,
                CommandBinding,
                ExecutedRoutedEventHandler,
                Key,
                KeyBinding,
                ModifierKeys,
                MouseButtonEventHandler,
            )

            self._win.CommandBindings.Add(
                CommandBinding(
                    ApplicationCommands.Close,
                    ExecutedRoutedEventHandler(
                        lambda s, e: self._on_cancel(self._win, None)
                    ),
                )
            )
            self._win.InputBindings.Add(
                KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None)
            )
            btn_close = self._win.FindName(u"BtnClose")
            title_bar = self._win.FindName(u"TitleBar")
            if title_bar is not None:
                title_bar.MouseLeftButtonDown += MouseButtonEventHandler(
                    lambda s, e: self._win.DragMove()
                )
            if btn_close is not None:
                btn_close.Click += RoutedEventHandler(self._on_cancel)
                btn_close.MouseLeftButtonDown += MouseButtonEventHandler(
                    lambda s, e: setattr(e, "Handled", True)
                )
        except Exception:
            pass
        self._grid.CellEditEnding += EventHandler[DataGridCellEditEndingEventArgs](
            self._on_receta_edit_end
        )
        self._grid.Loaded += RoutedEventHandler(self._on_grid_loaded)
        self._win.Loaded += RoutedEventHandler(self._on_win_loaded_opciones)
        if self._bd_opciones is not None:
            self._bd_opciones.SizeChanged += SizeChangedEventHandler(self._on_opciones_host_size)

        if owner_wpf is not None:
            try:
                from System.Windows.Interop import WindowInteropHelper

                oh = WindowInteropHelper(owner_wpf).Handle
                try:
                    if oh.ToInt64() != 0:
                        WindowInteropHelper(self._win).Owner = oh
                except Exception:
                    WindowInteropHelper(self._win).Owner = oh
            except Exception:
                pass

        self._refresh_preview()

    def _sync_lst_opciones_max_height(self):
        """Acota la altura del ListBox al marco para activar scroll interno."""
        try:
            bd = self._bd_opciones
            lb = self._lst_opciones
            if bd is None or lb is None:
                return
            h = float(bd.ActualHeight)
            if h > 24.0:
                lb.MaxHeight = max(h - 4.0, 100.0)
        except Exception:
            pass

    def _on_win_loaded_opciones(self, sender, args):
        self._sync_lst_opciones_max_height()

    def _on_opciones_host_size(self, sender, args):
        self._sync_lst_opciones_max_height()

    def _on_grid_loaded(self, sender, args):
        try:
            if self._grid.Columns.Count > 0:
                self._grid.Columns[self._grid.Columns.Count - 1].Visibility = Visibility.Collapsed
        except Exception:
            pass

    def _on_filtro_opciones(self, sender, args):
        dv = self._opts_dt.DefaultView
        try:
            t = self._txt_filtro.Text
            t = unicode(t).strip() if t is not None else u""
        except Exception:
            t = u""
        if not t:
            dv.RowFilter = u""
            return
        esc = t.replace(u"'", u"''")
        dv.RowFilter = u"[Label] LIKE '%{0}%'".format(esc)

    def _on_anadir(self, sender, args):
        try:
            sel = self._lst_opciones.SelectedItem
            if sel is None:
                return
            key = unicode(sel.Row[u"Key"])
            lab = unicode(sel.Row[u"Label"])
        except Exception:
            return
        for i in range(self._recipe.Rows.Count):
            try:
                if unicode(self._recipe.Rows[i][u"Key"]) == key:
                    return
            except Exception:
                pass
        row = self._recipe.NewRow()
        row[u"Key"] = key
        row[u"Origen"] = lab
        row[u"Prefix"] = u""
        row[u"Suffix"] = u""
        row[u"Separator"] = u"_"
        self._recipe.Rows.Add(row)
        self._refresh_preview()

    def _on_quitar(self, sender, args):
        try:
            idx = self._grid.SelectedIndex
        except Exception:
            idx = -1
        if idx < 0 or idx >= self._recipe.Rows.Count:
            return
        self._recipe.Rows.RemoveAt(idx)
        self._refresh_preview()

    def _on_subir(self, sender, args):
        idx = self._grid.SelectedIndex
        if idx <= 0:
            return
        _swap_recipe_rows(self._recipe, idx, idx - 1)
        self._grid.SelectedIndex = idx - 1
        self._refresh_preview()

    def _on_bajar(self, sender, args):
        idx = self._grid.SelectedIndex
        if idx < 0 or idx >= self._recipe.Rows.Count - 1:
            return
        _swap_recipe_rows(self._recipe, idx, idx + 1)
        self._grid.SelectedIndex = idx + 1
        self._refresh_preview()

    def _on_limpiar(self, sender, args):
        try:
            while self._recipe.Rows.Count > 0:
                self._recipe.Rows.RemoveAt(self._recipe.Rows.Count - 1)
        except Exception:
            pass
        self._refresh_preview()

    def _on_receta_edit_end(self, sender, args):
        self._refresh_preview()

    def _on_ok(self, sender, args):
        from Autodesk.Revit.UI import TaskDialog

        try:
            self._grid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

        segs = _recipe_segments_from_table(self._recipe)
        if not segs:
            TaskDialog.Show(
                u"Nombre Personalizado",
                u"Añada al menos un parámetro a la secuencia (flecha derecha).",
            )
            return

        n_ok = 0
        for i in range(self._main.Rows.Count):
            row = self._main.Rows[i]
            try:
                sid = int(row[u"IdInt"])
            except Exception:
                continue
            try:
                el = self._doc.GetElement(ElementId(sid))
            except Exception:
                continue
            if not isinstance(el, ViewSheet):
                continue
            try:
                name = self._evaluate(el, self._doc, segs)
                row[u"CustomName"] = name
                n_ok += 1
            except Exception:
                continue

        if n_ok == 0:
            TaskDialog.Show(
                u"Nombre Personalizado",
                u"No se actualizó ninguna fila. No hay láminas en el proyecto o no se pudieron leer como ViewSheet.",
            )
            return

        try:
            import exportar_laminas_pdf_dwg as _ex_nm  # noqa: WPS433

            _ex_nm.persist_naming_recipe_segments(self._doc, segs)
        except Exception:
            pass

        try:
            self._win.DialogResult = True
        except Exception:
            try:
                self._win.Close()
            except Exception:
                pass

    def _on_cancel(self, sender, args):
        try:
            self._win.DialogResult = False
        except Exception:
            try:
                self._win.Close()
            except Exception:
                pass

    def _refresh_preview(self):
        sh = _sample_sheet(self._doc, self._main)
        segs = _recipe_segments_from_table(self._recipe)
        if sh is None:
            self._txt_prev.Text = u"Sin lámina de referencia en el proyecto."
            return
        if not segs:
            self._txt_prev.Text = u"Defina la secuencia con la flecha \u2192."
            return
        try:
            prev = self._evaluate(sh, self._doc, segs)
        except Exception as ex:
            prev = u"(error: {0})".format(unicode(str(ex)))
        try:
            sn = sh.SheetNumber or u""
        except Exception:
            sn = u""
        self._txt_prev.Text = u"{0}\n\u2014 Ref.: {1}".format(prev, sn)

    def show_modal(self):
        try:
            self._win.ShowDialog()
        except Exception:
            pass


def show_componer_nombre_dialog(owner_wpf, doc, main_datatable, list_options_fn, evaluate_fn):
    """
    Muestra el diálogo Nombre Personalizado. Al confirmar, escribe CustomName en main_datatable.
    """
    dlg = ComponerNombreLaminaDialog(owner_wpf, doc, main_datatable, list_options_fn, evaluate_fn)
    dlg.show_modal()
