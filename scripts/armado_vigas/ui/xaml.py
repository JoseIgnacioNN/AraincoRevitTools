# -*- coding: utf-8 -*-
"""Plantilla XAML — ventana principal Armado vigas (tema oscuro BIMTools)."""

from armado_vigas.ui import layout as lay

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""

XAML_ARMADO_VIGAS = u"""<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  xmlns:po="http://schemas.microsoft.com/winfx/2006/xaml/presentation/options"
  Title="Arainco"
  Height="920" Width="1360"
  MinHeight="640" MinWidth="960"
  ResizeMode="CanResize"
  WindowStartupLocation="Manual"
  Background="#071018"
  FontFamily="Segoe UI"
  FontSize="12"
  ShowInTaskbar="False">
  <Window.Resources>
__BIMTOOLS_DARK_STYLES__
    <!-- Brushes estáticos del shell (Freezables). Estilos del tema dark no se modifican. -->
    <SolidColorBrush x:Key="ArmadoAppBg" Color="#071018" po:Freeze="True"/>
    <SolidColorBrush x:Key="ArmadoPanelBg" Color="#0a1620" po:Freeze="True"/>
    <SolidColorBrush x:Key="ArmadoBorder" Color="#21465C" po:Freeze="True"/>
    <SolidColorBrush x:Key="ArmadoFgHi" Color="#E8F4F8" po:Freeze="True"/>
    <SolidColorBrush x:Key="ArmadoFgMid" Color="#95B8CC" po:Freeze="True"/>
    <SolidColorBrush x:Key="ArmadoFgLo" Color="#64748b" po:Freeze="True"/>
    <SolidColorBrush x:Key="ArmadoAccentSoft" Color="#7eb8d0" po:Freeze="True"/>
  </Window.Resources>
  <Border Background="{StaticResource ArmadoAppBg}" BorderBrush="{StaticResource ArmadoBorder}" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Cabecera en Grid (menos anidamiento que StackPanel). -->
      <Grid Grid.Row="0" Margin="0,0,0,8">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="TxtTitle" Grid.Row="0" Text="Arainco: Armado vigas"
                   Foreground="{StaticResource ArmadoFgHi}" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Grid.Row="1" Margin="0,6,0,0"
                   Foreground="{StaticResource ArmadoFgMid}" TextWrapping="Wrap"
                   Text="Vista previa UI · alzado + sección · pestañas SUP / INF / LAT / CONF (sin colocación de Rebar)"/>
      </Grid>

      <Border Grid.Row="1" Margin="0,0,0,8" Background="{StaticResource ArmadoPanelBg}"
              BorderBrush="{StaticResource ArmadoBorder}"
              BorderThickness="1" CornerRadius="4" Padding="8,6">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock x:Name="TxtTramoSummary" Grid.Row="0"
                     Foreground="{StaticResource ArmadoAccentSoft}" FontSize="10" FontWeight="SemiBold"/>
          <TextBlock x:Name="TxtApoyosSummary" Grid.Row="1" Margin="0,4,0,0"
                     Foreground="{StaticResource ArmadoFgLo}" FontSize="10"/>
        </Grid>
      </Border>

      <TextBlock x:Name="TxtSelectionInfo" Grid.Row="2"
                 Foreground="{StaticResource ArmadoFgLo}" FontSize="10"
                 TextWrapping="Wrap" Margin="0,0,0,10"
                 Text="Clic viga → selección · Ctrl+clic pill Tn → multi · rueda pan H · MMB arrastrar · Ctrl+rueda zoom · Emp/Tn alzado · rail SUP/INF."/>

      <Grid Grid.Row="3">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="__SECTION_RAIL_WIDTH__"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="{StaticResource ArmadoPanelBg}"
                BorderBrush="{StaticResource ArmadoBorder}" BorderThickness="1"
                CornerRadius="4,0,0,4" Padding="0">
          <!-- Canvas host imperativo: virtualización de ítems no aplica; deferred scroll + virtualizing stack alivia layout. -->
          <ScrollViewer x:Name="ScrCanvas"
                        VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Auto"
                        CanContentScroll="False"
                        IsDeferredScrollingEnabled="True"
                        RenderOptions.EdgeMode="Aliased"
                        SnapsToDevicePixels="True">
            <Grid x:Name="PnlCanvasHost" Background="Transparent" SnapsToDevicePixels="True"
                  RenderOptions.EdgeMode="Aliased"
                  VerticalAlignment="Top" HorizontalAlignment="Left"/>
          </ScrollViewer>
        </Border>

        <Border Grid.Column="1" Background="{StaticResource ArmadoPanelBg}"
                BorderBrush="{StaticResource ArmadoBorder}" BorderThickness="1,1,1,1"
                CornerRadius="0,4,4,0" Padding="8,8">
          <ScrollViewer VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Disabled"
                        CanContentScroll="False"
                        IsDeferredScrollingEnabled="True">
            <!-- Controles imperativos en code-behind; Grid raíz reduce niveles.
                 Virtualización ItemsControl: el rail se genera en Python (PnlSectionCtrls).
                 Cuando se migre a ListBox/DataGrid, añadir:
                 VirtualizingStackPanel.IsVirtualizing=True + VirtualizationMode=Recycling. -->
            <Grid x:Name="PnlSectionRail" SnapsToDevicePixels="True">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Grid x:Name="GrdSectionRailHint" Grid.Row="0" Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="TxtSectionRailHint" Grid.Column="0"
                           Text="Sección · confinamiento"
                           Foreground="{StaticResource ArmadoFgLo}" FontSize="9"
                           VerticalAlignment="Center" TextWrapping="Wrap"
                           Margin="0,0,8,0"/>
                <Button x:Name="BtnSectionZoom" Grid.Column="1" Content="Zoom"
                        Padding="8,3" FontSize="10" Cursor="Hand"
                        Visibility="Collapsed"
                        ToolTip="Abrir sección ampliada (CONF) para dibujar estribos/trabas"
                        Style="{StaticResource BtnSelectOutline}" MinWidth="56"/>
              </Grid>
              <Border x:Name="BdrSectionPreview" Grid.Row="1" Background="{StaticResource ArmadoAppBg}"
                      BorderBrush="{StaticResource ArmadoBorder}"
                      BorderThickness="1" CornerRadius="4" Padding="2" MinHeight="236"
                      SnapsToDevicePixels="True">
                <Canvas x:Name="CnvSectionPreview" Width="__PREVIEW_CANVAS_W__" Height="222"
                        SnapsToDevicePixels="True"
                        RenderOptions.EdgeMode="Aliased"/>
              </Border>
              <TextBlock x:Name="TxtSectionMeta" Grid.Row="2" Margin="0,8,0,0"
                         Foreground="{StaticResource ArmadoFgMid}"
                         FontSize="10" TextWrapping="Wrap"/>
              <StackPanel x:Name="PnlSectionCtrls" Grid.Row="3" Margin="0,8,0,0"
                          HorizontalAlignment="Stretch"/>
            </Grid>
          </ScrollViewer>
        </Border>
      </Grid>

      <Grid Grid.Row="4" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="TxtEstado" Grid.Column="0" VerticalAlignment="Center"
                   Foreground="{StaticResource ArmadoFgLo}" FontSize="10"
                   TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnCancelar" Content="Cerrar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnColocar" Content="Colocar armadura"
                  Style="{StaticResource BtnPrimary}" MinWidth="200" IsEnabled="True"
                  ToolTip="Crea Rebar según toggles SUP / INF / LAT / CONF del rail"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>"""


def build_armado_vigas_xaml():
    xaml = XAML_ARMADO_VIGAS.replace(u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)
    xaml = xaml.replace(
        u"__SECTION_RAIL_WIDTH__",
        u"{0:.0f}".format(lay.SECTION_RAIL_WIDTH_PX),
    )
    xaml = xaml.replace(
        u"__PREVIEW_CANVAS_W__",
        u"{0:.0f}".format(lay.SECTION_CTRL_WIDTH_PX),
    )
    return xaml
