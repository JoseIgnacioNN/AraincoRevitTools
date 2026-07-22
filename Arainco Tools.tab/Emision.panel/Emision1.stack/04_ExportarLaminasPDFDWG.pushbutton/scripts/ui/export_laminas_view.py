# -*- coding: utf-8 -*-
"""
Vista de la herramienta Exportar Láminas.

ExportarLaminasView envuelve la ventana WPF (XAML, controles nombrados,
listeners de eventos) y delega toda la lógica al ViewModel mediante llamadas
directas a métodos y comandos. No contiene lógica de negocio.

La Vista:
  – Parsea el XAML y localiza todos los controles nombrados.
  – Registra los callbacks de UI en el ViewModel (Observer ligero).
  – Conecta eventos WPF a los métodos del ViewModel.
  – Aplica utilidades de scrollbar y logo propias del proyecto BIMTools.
"""

import clr  # noqa: E402

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Data")

from System import Action, EventHandler  # noqa: E402
from System.Data import DataRowChangeEventHandler  # noqa: E402
from System.Windows import RoutedEventHandler, Visibility  # noqa: E402
from System.Windows.Controls import DataGridCellEditEndingEventArgs  # noqa: E402
from System.Windows.Markup import XamlReader  # noqa: E402
import System  # noqa: E402

try:
    from infra.bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML as _DARK_STYLES
except Exception:
    _DARK_STYLES = u""

# ---------------------------------------------------------------------------
# XAML de la ventana
# ---------------------------------------------------------------------------

_XAML = (
    u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    x:Name="ExpLamWin"
    Title="Arainco: Exportar Láminas"
    Height="810" Width="1040" MinHeight="640" MinWidth="800"
    ResizeMode="CanResize"
    WindowStartupLocation="Manual"
    Background="#071018"
    FontFamily="Segoe UI" FontSize="12"
    ShowInTaskbar="False">
  <Window.Resources>
"""
    + _DARK_STYLES
    + u"""
    <Style TargetType="ComboBoxItem" BasedOn="{StaticResource ComboItem}"/>
    <Style x:Key="ExpLamFechaCombo" TargetType="ComboBox" BasedOn="{StaticResource ComboStretch}">
      <Setter Property="MinHeight" Value="32"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
    </Style>
    <Style x:Key="ExpFmtToggle" TargetType="CheckBox">
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Border x:Name="Bd" Background="#0A1627" BorderBrush="#5BC0DE" BorderThickness="1" CornerRadius="5" Padding="8,5" MinHeight="30" HorizontalAlignment="Stretch" SnapsToDevicePixels="True">
              <Grid HorizontalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Box" Grid.Column="0" Width="14" Height="14" Background="#050E18" BorderBrush="#21465C" BorderThickness="1" CornerRadius="2" Margin="0,0,6,0" VerticalAlignment="Center">
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
                <Setter TargetName="Bd" Property="Background" Value="#0E1B32"/>
                <Setter TargetName="Bd" Property="BorderBrush" Value="#7ED4ED"/>
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
      <Setter Property="Background" Value="#0a1620"/>
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="RowBackground" Value="#0B1726"/>
      <Setter Property="AlternatingRowBackground" Value="#071018"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#21465C"/>
      <Setter Property="VerticalGridLinesBrush" Value="#21465C"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="RowHeight" Value="34"/>
      <Setter Property="GridLinesVisibility" Value="All"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource {x:Type DataGridColumnHeader}}">
      <Setter Property="Background" Value="#11253D"/>
      <Setter Property="Foreground" Value="#95B8CC"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Padding" Value="12,10"/>
      <Setter Property="BorderBrush" Value="#21465C"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>
    <Style TargetType="DataGridRow" BasedOn="{StaticResource {x:Type DataGridRow}}">
      <Setter Property="Background" Value="#0B1726"/>
      <Style.Triggers>
        <Trigger Property="AlternationIndex" Value="0">
          <Setter Property="Background" Value="#0B1726"/>
        </Trigger>
        <Trigger Property="AlternationIndex" Value="1">
          <Setter Property="Background" Value="#071018"/>
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
    <Style x:Key="DgEditLeft" TargetType="TextBox" BasedOn="{StaticResource BimToolsTextBoxDark}">
      <Setter Property="TextAlignment" Value="Left"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="FontSize" Value="12"/>
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
    <Style x:Key="DgTbCenter" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#E8F4F8"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
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
    <Style x:Key="BtnNombreArchivoHeader" TargetType="Button" BasedOn="{StaticResource BtnSelectOutline}">
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Padding" Value="12,10"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="DgHdrNombreArchivo" TargetType="DataGridColumnHeader" BasedOn="{StaticResource DgHdrLeft}">
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Padding" Value="4,6"/>
      <Setter Property="Background" Value="Transparent"/>
    </Style>
    <!-- Misma barra que BIMTools (flechas rellenas, ~18px): la versión ExpLam 10px + trazos se leía mal. -->
    <Style x:Key="ExpLamScrollBarDark" TargetType="ScrollBar" BasedOn="{StaticResource BimToolsScrollBarDark}"/>
  </Window.Resources>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1" Padding="18">
    <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <StackPanel Grid.Row="0" Margin="0,0,0,10">
      <TextBlock x:Name="TxtTitle" Text="Arainco: Exportar Láminas" Foreground="#E8F4F8" FontSize="18" FontWeight="Bold"/>
      <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0" Foreground="#95B8CC" TextWrapping="Wrap"
                 Text="Selecciona láminas y exporta a PDF, DWG y/o listado Excel."/>
    </StackPanel>
    <Border Grid.Row="1" Background="#0a1620" BorderBrush="#21465C" BorderThickness="1"
            CornerRadius="4" Padding="10,8" Margin="0,0,0,10">
      <Grid VerticalAlignment="Center">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid Grid.Column="0" Margin="0,0,10,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Border Grid.Column="0" Margin="0,0,6,0" Background="#050E18" BorderBrush="#21465C" BorderThickness="1" CornerRadius="4" Padding="0" MinHeight="32">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="&#xE721;" FontFamily="Segoe MDL2 Assets" FontSize="15"
                         Foreground="#7AA3B8" VerticalAlignment="Center" Margin="10,0,4,0" IsHitTestVisible="False"/>
              <Grid Grid.Column="1" MinHeight="28">
                <TextBox x:Name="TxtBuscar" Background="Transparent" BorderThickness="0"
                         Foreground="#E8F4F8" CaretBrush="#7AA3B8"
                         VerticalContentAlignment="Center" Padding="0,6,10,6"
                         ToolTip="Filtrar por número, nombre o fecha de entrega (lista principal)"/>
                <TextBlock x:Name="TxtBuscarWatermark" Text="Buscar" IsHitTestVisible="False"
                           Foreground="#64748b" FontSize="12" VerticalAlignment="Center" Margin="0,0,10,0"/>
              </Grid>
            </Grid>
          </Border>
          <Border Grid.Column="1" Background="#050E18" BorderBrush="#21465C" BorderThickness="1" CornerRadius="4"
                  Padding="0" MinHeight="32">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="&#xE787;" FontFamily="Segoe MDL2 Assets" FontSize="15"
                         Foreground="#7AA3B8" VerticalAlignment="Center" Margin="10,0,4,0" IsHitTestVisible="False"
                         ToolTip="Fecha de entrega (parámetros FCH)"/>
              <Grid Grid.Column="1" MinHeight="28">
                <ComboBox x:Name="CmbFechaEntrega" Style="{StaticResource ExpLamFechaCombo}"
                          VerticalAlignment="Center" VerticalContentAlignment="Center"
                          ToolTip="Al elegir una fecha se marcan automáticamente los planos con esa fecha en cualquier parámetro FCH."/>
                <TextBlock x:Name="TxtFechaEntregaWatermark" Text="Fecha de entrega" IsHitTestVisible="False"
                           Foreground="#64748b" FontSize="12" VerticalAlignment="Center" Margin="6,0,10,0"/>
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
        <Button x:Name="BtnRefrescar" Grid.Column="4" Style="{StaticResource BtnSelectOutline}" MinWidth="110"
                ToolTip="Actualizar la lista de láminas desde el proyecto">
          <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
            <TextBlock Text="&#xE72C;" FontFamily="Segoe MDL2 Assets" FontSize="13"
                       VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBlock Text="Actualizar" VerticalAlignment="Center" FontSize="11" FontWeight="SemiBold"/>
          </StackPanel>
        </Button>
      </Grid>
    </Border>
    <Border Grid.Row="2" Background="#0a1620" BorderBrush="#21465C" BorderThickness="1" CornerRadius="4" Padding="0" Margin="0,0,0,0">
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
    <TextBlock Grid.Row="3" Text="Carpeta de salida" Style="{StaticResource Label}" Margin="0,12,0,8"/>
    <Border Grid.Row="4" Background="#0a1620" BorderBrush="#21465C" BorderThickness="1"
            CornerRadius="4" Padding="10,8" Margin="0,0,0,10">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="TxtCarpeta" Grid.Column="0" Style="{StaticResource BimToolsTextBoxDark}"
                 MinHeight="32" FontSize="12" Margin="0,0,12,0"/>
        <Button x:Name="BtnCarpeta" Grid.Column="1" Content="Examinar…" Style="{StaticResource BtnSelectOutline}"
                MinWidth="132"/>
      </Grid>
    </Border>
    <Grid Grid.Row="5" Margin="0,14,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBlock x:Name="TxtEstado" Grid.Column="0" VerticalAlignment="Center"
                 Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,12,0"/>
      <Button x:Name="BtnExportar" Grid.Column="1" Content="Exportar" Style="{StaticResource BtnPrimary}" MinWidth="132"/>
    </Grid>
    </Grid>
  </Border>
</Window>
"""
)


# ---------------------------------------------------------------------------
# Utilidades de Vista
# ---------------------------------------------------------------------------

def _apply_scrollbar_styles(root_visual, resources_owner):
    """Fuerza estilo oscuro BIMTools en todos los ScrollBar del árbol visual."""
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


def _schedule_scrollbar_styles(win, grid):
    """Aplica estilos de scrollbar en varias pasadas (el DataGrid regenera el ScrollViewer)."""
    try:
        from System.Windows.Threading import DispatcherPriority

        def _go():
            _apply_scrollbar_styles(grid, win)
            _apply_scrollbar_styles(win, win)

        _go()
        win.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_go))
        win.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, Action(_go))
    except Exception:
        try:
            _apply_scrollbar_styles(grid, win)
            _apply_scrollbar_styles(win, win)
        except Exception:
            pass


def _find_visual_child(root, predicate, max_depth=60):
    """Recorre el árbol visual en profundidad hasta encontrar el primer elemento que cumple predicate."""
    from System.Windows.Media import VisualTreeHelper

    def _walk(o, depth):
        if depth > max_depth or o is None:
            return None
        try:
            n = VisualTreeHelper.GetChildrenCount(o)
        except Exception:
            return None
        for i in range(n):
            try:
                ch = VisualTreeHelper.GetChild(o, i)
            except Exception:
                continue
            try:
                if predicate(ch):
                    return ch
            except Exception:
                pass
            found = _walk(ch, depth + 1)
            if found is not None:
                return found
        return None

    return _walk(root, 0)


# ---------------------------------------------------------------------------
# Vista
# ---------------------------------------------------------------------------

class ExportarLaminasView(object):
    """
    Vista WPF de la herramienta Exportar Láminas.

    Responsabilidades:
      – Parsear el XAML y localizar controles nombrados.
      – Registrar callbacks de UI en el ViewModel.
      – Conectar eventos WPF (Click, TextChanged, etc.) al ViewModel.
      – Aplicar estilo de scrollbar y logo.
      – Mostrar la ventana con ShowDialog().

    No contiene lógica de negocio.
    """

    # Referencia de clase: ancla fuerte en IronPython para prevenir GC
    # después de que main() retorna.
    _live_view_ref = None

    def __init__(
        self,
        view_model,
        folder_svc,
        revit_svc,
        show_componer_nombre_fn,
        appdomain_win_key,
    ):
        self._vm = view_model
        self._folder_svc = folder_svc
        self._revit_svc = revit_svc
        self._show_componer_fn = show_componer_nombre_fn
        self._appdomain_win_key = appdomain_win_key

        # Estado de vista (no de negocio)
        self._syncing_select_all = False
        self._chk_select_all = None
        self._sel_anchor_idx = None

        # Parsear ventana y encontrar controles
        self._win = XamlReader.Parse(_XAML)
        self._find_controls()
        self._wire_vm_callbacks()
        self._wire_events()

        # Poblar datos iniciales
        self._grid.ItemsSource = self._vm.table_view
        self._refresh_fch_combo()
        self._refresh_estado(self._vm.get_estado_text())

    # -----------------------------------------------------------------------
    # Controles
    # -----------------------------------------------------------------------

    def _find_controls(self):
        w = self._win
        self._grid = w.FindName("GridSheets")
        self._txt_buscar = w.FindName("TxtBuscar")
        self._txt_buscar_watermark = w.FindName("TxtBuscarWatermark")
        self._txt_carpeta = w.FindName("TxtCarpeta")
        self._txt_estado = w.FindName("TxtEstado")
        self._chk_pdf = w.FindName("ChkPdf")
        self._chk_dwg = w.FindName("ChkDwg")
        self._chk_listado_plan = w.FindName("ChkListadoPlanos")
        self._cmb_fecha_entrega = w.FindName("CmbFechaEntrega")
        self._txt_fecha_entrega_watermark = w.FindName("TxtFechaEntregaWatermark")
        self._btn_export = w.FindName("BtnExportar")
        self._btn_carpeta = w.FindName("BtnCarpeta")
        self._suppress_fch_changed = False

    # -----------------------------------------------------------------------
    # Registro de callbacks en el ViewModel
    # -----------------------------------------------------------------------

    def _wire_vm_callbacks(self):
        vm = self._vm

        vm.bind_on_show_ok(
            lambda msg: self._revit_svc.show_ok(msg, self._win)
        )
        vm.bind_on_show_errors(
            lambda instr, errs: self._revit_svc.show_errors(instr, errs, self._win)
        )
        vm.bind_on_ask_open_folder(
            lambda folder: self._revit_svc.ask_yes_no(
                u"¿Desea abrir la carpeta con todos los archivos generados?",
                folder,
                self._win,
            )
        )
        vm.bind_on_exporting_changed(self._on_exporting_changed)
        vm.bind_on_estado_changed(self._refresh_estado)
        vm.bind_get_fecha_emision(self._get_fecha_emision)

    # -----------------------------------------------------------------------
    # Conexión de eventos WPF
    # -----------------------------------------------------------------------

    def _wire_events(self):
        vm = self._vm

        # DataTable → estado (DataRowChangeEventHandler, no EventHandler genérico)
        vm.table.RowChanged += DataRowChangeEventHandler(lambda s, e: vm.on_row_changed())

        # DataGrid
        self._grid.Loaded += RoutedEventHandler(self._on_grid_loaded)
        self._grid.CellEditEnding += EventHandler[DataGridCellEditEndingEventArgs](
            lambda s, e: vm.on_cell_edit_ending()
        )
        try:
            from System.Windows.Input import MouseButtonEventHandler
            self._grid.PreviewMouseLeftButtonDown += MouseButtonEventHandler(
                self._on_grid_preview_mouse_left
            )
        except Exception:
            pass

        # Buscar (TextChangedEventHandler – no RoutedEventHandler)
        self._txt_buscar.TextChanged += self._on_buscar_changed
        self._txt_buscar.GotFocus += RoutedEventHandler(
            lambda s, e: self._sync_buscar_watermark()
        )
        self._txt_buscar.LostFocus += RoutedEventHandler(
            lambda s, e: self._sync_buscar_watermark()
        )
        self._sync_buscar_watermark()

        # TxtCarpeta: sincroniza carpeta del VM en tiempo real (TextChangedEventHandler)
        self._txt_carpeta.TextChanged += self._on_txt_carpeta_changed_handler

        # Formato checkboxes (ChkPdf, ChkDwg, ChkListadoPlanos)
        self._chk_pdf.Click += RoutedEventHandler(
            lambda s, e: setattr(vm, "do_pdf", self._nullable_bool(self._chk_pdf.IsChecked))
        )
        self._chk_dwg.Click += RoutedEventHandler(
            lambda s, e: setattr(vm, "do_dwg", self._nullable_bool(self._chk_dwg.IsChecked))
        )
        if self._chk_listado_plan is not None:
            self._chk_listado_plan.Click += RoutedEventHandler(
                lambda s, e: setattr(
                    vm, "do_listado",
                    self._nullable_bool(self._chk_listado_plan.IsChecked)
                )
            )

        # Botones principales
        self._win.FindName("BtnRefrescar").Click += RoutedEventHandler(self._on_refrescar)
        self._btn_carpeta.Click += RoutedEventHandler(self._on_carpeta_click)
        self._btn_export.Click += RoutedEventHandler(self._on_export_click)

        # Combo fecha de entrega
        try:
            from System.Windows.Controls import SelectionChangedEventHandler
            if self._cmb_fecha_entrega is not None:
                self._cmb_fecha_entrega.SelectionChanged += SelectionChangedEventHandler(
                    self._on_fch_selection_changed
                )
                self._cmb_fecha_entrega.GotFocus += RoutedEventHandler(
                    lambda s, e: self._sync_fch_watermark()
                )
                self._cmb_fecha_entrega.LostFocus += RoutedEventHandler(
                    lambda s, e: self._sync_fch_watermark()
                )
                try:
                    self._cmb_fecha_entrega.DropDownOpened += RoutedEventHandler(
                        lambda s, e: self._sync_fch_watermark()
                    )
                    self._cmb_fecha_entrega.DropDownClosed += RoutedEventHandler(
                        lambda s, e: self._sync_fch_watermark()
                    )
                except Exception:
                    pass
        except Exception:
            pass

        # Ventana: Escape para cerrar
        self._wire_close_keyboard_shortcut()

        self._win.Loaded += RoutedEventHandler(self._on_win_loaded)
        self._win.Closed += EventHandler(self._on_win_closed)

    def _wire_close_keyboard_shortcut(self):
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
                    ExecutedRoutedEventHandler(lambda s, e: self._win.Close()),
                )
            )
            self._win.InputBindings.Add(
                KeyBinding(ApplicationCommands.Close, Key.Escape, ModifierKeys.None)
            )
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Handlers de eventos
    # -----------------------------------------------------------------------

    def _on_win_loaded(self, sender, args):
        try:
            _schedule_scrollbar_styles(self._win, self._grid)
        except Exception:
            pass

    def _on_win_closed(self, sender, args):
        try:
            ExportarLaminasView._live_view_ref = None
        except Exception:
            pass
        try:
            System.AppDomain.CurrentDomain.SetData(self._appdomain_win_key, None)
        except Exception:
            pass

    def _on_grid_loaded(self, sender, args):
        try:
            if self._grid.Columns.Count > 0:
                self._grid.Columns[
                    self._grid.Columns.Count - 1
                ].Visibility = Visibility.Collapsed
        except Exception:
            pass
        try:
            _schedule_scrollbar_styles(self._win, self._grid)
        except Exception:
            pass
        try:
            from System.Windows.Threading import DispatcherPriority

            def _wire():
                try:
                    self._wire_custom_name_header()
                    self._wire_select_all_header()
                    self._sync_select_all_header()
                except Exception:
                    pass

            _wire()
            self._grid.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(_wire))
            self._grid.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_wire))
        except Exception:
            try:
                self._wire_custom_name_header()
                self._wire_select_all_header()
                self._sync_select_all_header()
            except Exception:
                pass

    def _on_buscar_changed(self, sender, args):
        self._sync_buscar_watermark()
        try:
            t = self._txt_buscar.Text
            t = unicode(t).strip() if t is not None else u""
        except Exception:
            t = u""
        self._vm.apply_search_filter(t)
        self._sync_select_all_header()

    def _on_txt_carpeta_changed_handler(self, sender, args):
        """TextChangedEventHandler para TxtCarpeta: sincroniza con el ViewModel."""
        try:
            self._vm.carpeta = self._txt_carpeta.Text
        except Exception:
            pass

    def _on_carpeta_click(self, sender, args):
        if getattr(self, "_carpeta_busy", False):
            return
        self._carpeta_busy = True
        try:
            cur = u""
            try:
                cur = unicode(self._txt_carpeta.Text).strip()
            except Exception:
                pass
            path = self._folder_svc.browse(cur, self._win)
            if path:
                self._vm.set_carpeta_from_browse(path)
                try:
                    self._txt_carpeta.Text = self._vm.carpeta
                except Exception:
                    pass
        finally:
            self._carpeta_busy = False

    def _on_export_click(self, sender, args):
        try:
            self._grid.CommitEdit()
        except Exception:
            pass
        self._vm.export_command.execute()

    def _on_refrescar(self, sender, args):
        self._chk_select_all = None
        self._sel_anchor_idx = None
        table, fch_names = self._vm.get_refreshed_table()
        self._grid.ItemsSource = table.DefaultView
        table.RowChanged += DataRowChangeEventHandler(lambda s, e: self._vm.on_row_changed())
        self._txt_buscar.Text = u""
        self._refresh_fch_combo()
        try:
            from System.Windows.Threading import DispatcherPriority

            def _wire():
                try:
                    self._wire_custom_name_header()
                    self._wire_select_all_header()
                    self._sync_select_all_header()
                except Exception:
                    pass

            self._grid.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(_wire))
            self._grid.Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, Action(_wire))
        except Exception:
            try:
                self._wire_custom_name_header()
                self._wire_select_all_header()
                self._sync_select_all_header()
            except Exception:
                pass

    def _on_fch_selection_changed(self, sender, args):
        if self._suppress_fch_changed:
            return
        cmb = self._cmb_fecha_entrega
        if cmb is None:
            return
        try:
            sel = cmb.SelectedItem
        except Exception:
            sel = None
        if sel is None:
            self._sync_fch_watermark()
            return
        try:
            key = unicode(sel).strip()
        except Exception:
            key = u""
        if key:
            self._vm.apply_fch_selection(key)
        self._sync_fch_watermark()

    def _on_exporting_changed(self, is_exporting):
        """Al iniciar exportación cierra el formulario; al terminar reactiva controles si sigue abierto."""
        if is_exporting:
            try:
                self._win.Close()
            except Exception:
                pass
            return
        enabled = True
        try:
            self._btn_export.IsEnabled = enabled
            self._chk_pdf.IsEnabled = enabled
            self._chk_dwg.IsEnabled = enabled
            if self._chk_listado_plan is not None:
                self._chk_listado_plan.IsEnabled = enabled
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Selección "Componer nombre" (encabezado de columna)
    # -----------------------------------------------------------------------

    def _wire_custom_name_header(self):
        from System.Windows.Controls import Button

        def pred(ch):
            try:
                return isinstance(ch, Button) and unicode(str(ch.Tag)) == u"HdrComponer"
            except Exception:
                return False

        btn = _find_visual_child(self._grid, pred)
        if btn is None:
            return
        if not hasattr(self, u"_hdr_componer_handler"):
            self._hdr_componer_handler = RoutedEventHandler(self._on_componer_nombre)
        try:
            btn.Click -= self._hdr_componer_handler
        except Exception:
            pass
        btn.Click += self._hdr_componer_handler

    def _on_componer_nombre(self, sender, args):
        try:
            args.Handled = True
        except Exception:
            pass
        try:
            self._show_componer_fn(
                self._win,
                self._vm.doc,
                self._vm.table,
                self._vm.list_naming_source_options,
                self._vm.evaluate_naming_recipe,
            )
        except Exception as ex:
            try:
                from ui.export_laminas_instruction_dialog import show_message_dialog
                from infra.revit_wpf_window_position import revit_main_hwnd

                revit = self._vm.revit
                uiapp = None
                try:
                    uiapp = revit.Application if revit is not None else None
                except Exception:
                    uiapp = None
                if uiapp is None:
                    uiapp = revit
                show_message_dialog(
                    u"Arainco: Nombre Personalizado",
                    u"Error al abrir Nombre Personalizado.",
                    unicode(str(ex)),
                    ok_text=u"Entendido",
                    hwnd_revit=revit_main_hwnd(uiapp),
                    uiapp=uiapp,
                )
            except Exception:
                pass
        self._refresh_estado(self._vm.get_estado_text())

    # -----------------------------------------------------------------------
    # Selección "Seleccionar todo" (cabecera de columna)
    # -----------------------------------------------------------------------

    def _wire_select_all_header(self):
        from System.Windows.Controls import CheckBox

        def pred(ch):
            try:
                return isinstance(ch, CheckBox) and unicode(str(ch.Tag)) == u"HdrSelectAll"
            except Exception:
                return False

        chk = _find_visual_child(self._grid, pred)
        if chk is None:
            return
        self._chk_select_all = chk
        if not hasattr(self, u"_select_all_handler"):
            self._select_all_handler = RoutedEventHandler(
                self._on_select_all_header_click
            )
        try:
            chk.Click -= self._select_all_handler
        except Exception:
            pass
        chk.Click += self._select_all_handler

    def _on_select_all_header_click(self, sender, args):
        if self._syncing_select_all:
            return
        try:
            args.Handled = True
        except Exception:
            pass
        self._syncing_select_all = True
        try:
            self._vm.toggle_all_visible()
        finally:
            self._syncing_select_all = False
        self._sync_select_all_header()

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
        n, n_sel = self._vm.get_visible_selection_state()
        self._syncing_select_all = True
        try:
            try:
                chk.IsThreeState = True
            except Exception:
                pass
            if n == 0:
                chk.IsChecked = False
            elif n_sel == 0:
                chk.IsChecked = False
            elif n_sel == n:
                chk.IsChecked = True
            else:
                chk.IsChecked = None
        finally:
            self._syncing_select_all = False

    # -----------------------------------------------------------------------
    # Click en checkbox del DataGrid (un clic alterna; Mayús+clic: rango)
    # -----------------------------------------------------------------------

    def _on_grid_preview_mouse_left(self, sender, e):
        try:
            from System.Windows.Controls import DataGridCell, DataGridCheckBoxColumn
            from System.Windows.Input import Keyboard, ModifierKeys
            from System.Windows.Media import VisualTreeHelper
            from System import Boolean
        except Exception:
            return

        # Encontrar la celda clickeada
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
        try:
            idx = g.Items.IndexOf(drv)
        except Exception:
            idx = -1
        if idx < 0:
            return

        def _sel_get(rv):
            try:
                return self._vm.row_is_selected(rv.Row)
            except Exception:
                try:
                    return self._vm._nullable_bool(rv[u"Sel"])
                except Exception:
                    return False

        shift_down = False
        try:
            shift_down = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
        except Exception:
            pass

        anchor = self._sel_anchor_idx

        if shift_down and anchor is not None:
            cur_on = _sel_get(drv)
            new_val = not cur_on
            i0 = min(int(anchor), int(idx))
            i1 = max(int(anchor), int(idx))
            row_views = []
            for j in range(i0, i1 + 1):
                try:
                    row_views.append(g.Items[j])
                except Exception:
                    pass
            self._vm.set_rows_selected(row_views, new_val)
            try:
                e.Handled = True
            except Exception:
                pass
            self._sync_select_all_header()
            return

        def _apply_bulk():
            cur_on = _sel_get(drv)
            new_val = not cur_on
            self._sel_anchor_idx = idx

            seen = set()
            indices = []
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

            row_views = []
            for j in indices:
                try:
                    row_views.append(g.Items[j])
                except Exception:
                    pass
            self._vm.set_rows_selected(row_views, new_val)
            self._sync_select_all_header()

        try:
            from System.Windows.Threading import DispatcherPriority
            g.Dispatcher.BeginInvoke(DispatcherPriority.Input, Action(_apply_bulk))
        except Exception:
            _apply_bulk()
        try:
            e.Handled = True
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Fecha de entrega
    # -----------------------------------------------------------------------

    def _refresh_fch_combo(self):
        cmb = self._cmb_fecha_entrega
        if cmb is None:
            return
        self._suppress_fch_changed = True
        try:
            try:
                cmb.Items.Clear()
            except Exception:
                pass
            fch_names = self._vm.fch_param_names
            if not fch_names:
                try:
                    cmb.IsEnabled = False
                    cmb.ToolTip = (
                        u"No hay parámetro de lámina cuyo nombre contenga «FCH». "
                        u"Se revisan todos los parámetros FCH por plano al elegir fecha."
                    )
                except Exception:
                    pass
            else:
                vals = self._vm.get_fch_unique_values()
                for v in vals:
                    try:
                        cmb.Items.Add(v)
                    except Exception:
                        pass
                try:
                    names_tip = u", ".join(fch_names[:8])
                    if len(fch_names) > 8:
                        names_tip += u"…"
                    cmb.IsEnabled = len(vals) > 0
                    cmb.ToolTip = (
                        u"Parámetros FCH: {0}. Al elegir un valor se marcan "
                        u"los planos que lo tienen."
                    ).format(names_tip)
                except Exception:
                    pass
        finally:
            self._suppress_fch_changed = False
        self._sync_fch_watermark()

    def _get_fecha_emision(self):
        """Devuelve la fecha seleccionada en CmbFechaEntrega (para el listado Excel)."""
        cmb = self._cmb_fecha_entrega
        if cmb is None:
            return None
        try:
            sel = cmb.SelectedItem
            if sel is None:
                return None
            key = unicode(sel).strip()
            return key if key else None
        except Exception:
            return None

    # -----------------------------------------------------------------------
    # Watermarks de Buscar y Fecha entrega
    # -----------------------------------------------------------------------

    def _sync_buscar_watermark(self):
        wm = self._txt_buscar_watermark
        tb = self._txt_buscar
        if wm is None or tb is None:
            return
        try:
            t = tb.Text
            t = unicode(t).strip() if t is not None else u""
        except Exception:
            t = u""
        focused = False
        try:
            focused = bool(tb.IsFocused)
        except Exception:
            pass
        show = (not t) and (not focused)
        try:
            wm.Visibility = Visibility.Visible if show else Visibility.Collapsed
        except Exception:
            pass

    def _sync_fch_watermark(self):
        wm = self._txt_fecha_entrega_watermark
        cmb = self._cmb_fecha_entrega
        if wm is None or cmb is None:
            return
        has_sel = False
        try:
            sel = cmb.SelectedItem
            if sel is not None:
                s = unicode(sel).strip()
                has_sel = bool(s)
        except Exception:
            pass
        dd_open = False
        try:
            dd_open = bool(cmb.IsDropDownOpen)
        except Exception:
            pass
        focused = False
        try:
            focused = bool(cmb.IsKeyboardFocusWithin)
        except Exception:
            try:
                focused = bool(cmb.IsFocused)
            except Exception:
                pass
        show = (not has_sel) and (not focused) and (not dd_open)
        try:
            wm.Visibility = Visibility.Visible if show else Visibility.Collapsed
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Estado
    # -----------------------------------------------------------------------

    def _refresh_estado(self, estado_text):
        try:
            self._txt_estado.Text = estado_text
        except Exception:
            pass
        if not getattr(self, "_syncing_select_all", False):
            self._sync_select_all_header()

    # -----------------------------------------------------------------------
    # Utilidades estáticas
    # -----------------------------------------------------------------------

    @staticmethod
    def _nullable_bool(wpf_nullable):
        try:
            if wpf_nullable is None:
                return False
            if hasattr(wpf_nullable, u"HasValue"):
                return bool(wpf_nullable.HasValue and wpf_nullable.Value)
            return unicode(wpf_nullable).strip().lower() == u"true"
        except Exception:
            return False

    # -----------------------------------------------------------------------
    # Mostrar la ventana
    # -----------------------------------------------------------------------

    def _attach_revit_owner(self):
        """Ventana hija de Revit (mismo patrón que Armado Muros)."""
        try:
            from System.Windows.Interop import WindowInteropHelper
            from infra.revit_wpf_window_position import revit_main_hwnd

            revit = self._vm.revit
            uiapp = None
            try:
                uiapp = revit.Application if revit is not None else None
            except Exception:
                uiapp = None
            if uiapp is None:
                uiapp = revit
            hwnd = revit_main_hwnd(uiapp)
            if hwnd is not None:
                try:
                    if hwnd.ToInt64() != 0:
                        WindowInteropHelper(self._win).Owner = hwnd
                except Exception:
                    WindowInteropHelper(self._win).Owner = hwnd
        except Exception:
            pass

    def _resolve_revit_hwnd(self):
        revit = self._vm.revit
        uiapp = None
        try:
            uiapp = revit.Application if revit is not None else None
        except Exception:
            uiapp = None
        if uiapp is None:
            uiapp = revit
        try:
            from infra.revit_wpf_window_position import revit_main_hwnd

            return revit_main_hwnd(uiapp)
        except Exception:
            return None

    def _prepare_window_bounds(self):
        """Sin tope MaxWidth/MaxHeight para que maximizar use el monitor de Revit."""
        try:
            from System import Double

            self._win.MaxWidth = Double.PositiveInfinity
            self._win.MaxHeight = Double.PositiveInfinity
        except Exception:
            pass

    def _position_window(self):
        """Centra el formulario en el monitor donde corre Revit."""
        try:
            from infra.revit_wpf_window_position import (
                bind_center_wpf_on_revit_monitor,
                bind_maximize_wpf_on_revit_monitor,
                position_wpf_window_center_on_monitor,
            )

            hwnd = self._resolve_revit_hwnd()
            bind_center_wpf_on_revit_monitor(self._win, hwnd)
            bind_maximize_wpf_on_revit_monitor(self._win, hwnd)
            position_wpf_window_center_on_monitor(self._win, hwnd)
        except Exception:
            pass

    def show(self):
        """
        Muestra la ventana con ShowDialog() (bucle de mensajes anidado que garantiza
        el routing de input ratón/teclado en el contexto Revit+pyRevit).
        """
        ExportarLaminasView._live_view_ref = self
        try:
            System.AppDomain.CurrentDomain.SetData(
                self._appdomain_win_key, self._win
            )
        except Exception:
            pass
        self._attach_revit_owner()
        self._prepare_window_bounds()
        self._position_window()
        try:
            self._win.ShowDialog()
        except Exception:
            pass
