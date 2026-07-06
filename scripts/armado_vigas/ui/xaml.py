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
  Title="Arainco: Armado vigas"
  Height="920" Width="1360"
  MinHeight="640" MinWidth="960" MaxWidth="1920"
  ResizeMode="CanResize"
  WindowStartupLocation="Manual"
  Background="#071018"
  FontFamily="Segoe UI"
  FontSize="12"
  ShowInTaskbar="False">
  <Window.Resources>
__BIMTOOLS_DARK_STYLES__
  </Window.Resources>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,8">
        <TextBlock x:Name="TxtTitle" Text="Arainco: Armado vigas"
                   Foreground="#E8F4F8" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0" Foreground="#95B8CC" TextWrapping="Wrap"
                   Text="Herramienta para el armado automático de vigas de hormigón armado"/>
      </StackPanel>

      <Border Grid.Row="1" Margin="0,0,0,8" Background="#0a1620" BorderBrush="#21465C"
              BorderThickness="1" CornerRadius="4" Padding="8,6">
        <StackPanel>
          <TextBlock x:Name="TxtTramoSummary" Foreground="#7eb8d0" FontSize="10" FontWeight="SemiBold"/>
          <TextBlock x:Name="TxtApoyosSummary" Margin="0,4,0,0" Foreground="#64748b" FontSize="10"/>
        </StackPanel>
      </Border>

      <TextBlock x:Name="TxtSelectionInfo" Grid.Row="2" Foreground="#64748b" FontSize="10"
                 TextWrapping="Wrap" Margin="0,0,0,10"
                 Text="Clic viga → selección · Traslape sup/inf en panel derecho · Controles Tn en bandas alzado · Cent/Ext L auto en panel sección."/>

      <Grid Grid.Row="3">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="__SECTION_RAIL_WIDTH__"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="#0a1620" BorderBrush="#21465C" BorderThickness="1"
                CornerRadius="4,0,0,4" Padding="0">
          <ScrollViewer x:Name="ScrCanvas" VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Auto">
            <StackPanel x:Name="PnlCanvasHost" Background="Transparent" SnapsToDevicePixels="True"/>
          </ScrollViewer>
        </Border>

        <Border Grid.Column="1" Background="#0a1620" BorderBrush="#21465C" BorderThickness="1,1,1,1"
                CornerRadius="0,4,4,0" Padding="8,8">
          <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <StackPanel x:Name="PnlSectionRail">
              <TextBlock x:Name="TxtSectionRailHint" Text="Sección · confinamiento"
                         Foreground="#64748b" FontSize="9" Margin="0,0,0,8"/>
              <Border x:Name="BdrSectionPreview" Background="#071018" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4" Padding="2" MinHeight="236">
                <Canvas x:Name="CnvSectionPreview" Width="__PREVIEW_CANVAS_W__" Height="222"/>
              </Border>
              <TextBlock x:Name="TxtSectionMeta" Margin="0,8,0,0" Foreground="#95B8CC"
                         FontSize="10" TextWrapping="Wrap"/>
              <StackPanel x:Name="PnlSectionCtrls" Margin="0,8,0,0"/>
            </StackPanel>
          </ScrollViewer>
        </Border>
      </Grid>

      <Grid Grid.Row="4" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="TxtEstado" Grid.Column="0" VerticalAlignment="Center"
                   Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnCancelar" Content="Cancelar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnColocar" Content="Colocar armadura"
                  Style="{StaticResource BtnPrimary}" MinWidth="180"/>
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
