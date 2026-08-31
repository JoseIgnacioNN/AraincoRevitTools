# -*- coding: utf-8 -*-
"""Plantilla XAML — Mallas en muros (shell alzado + rail, estilo Armado vigas)."""

from __future__ import print_function

try:
    from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
except Exception:
    BIMTOOLS_DARK_STYLES_XML = u""

try:
    from bimtools_ui_tokens import (
        ACCENT_PRIMARY,
        BG_APP,
        BG_PANEL,
        BORDER,
        FG_BODY,
        FG_MUTED,
        FG_TITLE,
        FONT_FAMILY,
        FONT_SIZE_BASE,
        WINDOW_CHROME_TITLE,
    )
except Exception:
    ACCENT_PRIMARY = u"#5BC0DE"
    BG_APP = u"#071018"
    BG_PANEL = u"#0a1620"
    BORDER = u"#21465C"
    FG_BODY = u"#95B8CC"
    FG_MUTED = u"#64748b"
    FG_TITLE = u"#E8F4F8"
    FONT_FAMILY = u"Segoe UI"
    FONT_SIZE_BASE = 12
    WINDOW_CHROME_TITLE = u"Arainco"

# Alineado con armado_vigas/ui/layout.py SECTION_RAIL_WIDTH_PX
MALLAS_SECTION_RAIL_WIDTH_PX = 340.0

XAML_MALLAS_EN_MUROS = u"""<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  xmlns:po="http://schemas.microsoft.com/winfx/2006/xaml/presentation/options"
  Title="__WINDOW_CHROME_TITLE__"
  Height="920" Width="1360"
  MinHeight="640" MinWidth="960"
  ResizeMode="CanResize"
  WindowStartupLocation="Manual"
  Background="__BG_APP__"
  FontFamily="__FONT_FAMILY__"
  FontSize="__FONT_SIZE_BASE__"
  ShowInTaskbar="False">
  <Window.Resources>
__BIMTOOLS_DARK_STYLES__
    <SolidColorBrush x:Key="MallasAppBg" Color="#071018" po:Freeze="True"/>
    <SolidColorBrush x:Key="MallasPanelBg" Color="#0a1620" po:Freeze="True"/>
    <SolidColorBrush x:Key="MallasBorder" Color="#21465C" po:Freeze="True"/>
    <SolidColorBrush x:Key="MallasFgHi" Color="#E8F4F8" po:Freeze="True"/>
    <SolidColorBrush x:Key="MallasFgMid" Color="#95B8CC" po:Freeze="True"/>
    <SolidColorBrush x:Key="MallasFgLo" Color="#64748b" po:Freeze="True"/>
    <SolidColorBrush x:Key="MallasAccent" Color="#7eb8d0" po:Freeze="True"/>
  </Window.Resources>
  <Border Background="{StaticResource MallasAppBg}" BorderBrush="{StaticResource MallasBorder}"
          BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Grid Grid.Row="0" Margin="0,0,0,8">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="TxtTitle" Grid.Row="0" Text="Arainco: Mallas en muros"
                   Foreground="{StaticResource MallasFgHi}" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Grid.Row="1" Margin="0,6,0,0"
                   Foreground="{StaticResource MallasFgMid}" TextWrapping="Wrap"
                   Text="Elevación a escala · Area Reinforcement ext.+int. (sin Inicio/Término)."/>
        <StackPanel x:Name="PnlPhaseStepper" Grid.Row="1" Orientation="Vertical"
                    Margin="0,0,0,0" Visibility="Collapsed"/>
        <StackPanel x:Name="PnlModoMuro" Grid.Row="1" Orientation="Horizontal"
                    Margin="0,8,0,0" Visibility="Collapsed">
          <CheckBox x:Name="ChkMuroTradicional" Content="Muro Tradicional" IsChecked="True"
                    Foreground="{StaticResource MallasFgHi}" FontSize="11" Margin="0,0,24,0"
                    VerticalAlignment="Center"/>
          <CheckBox x:Name="ChkMuroContencion" Content="Muro de Contención"
                    Foreground="{StaticResource MallasFgHi}" FontSize="11"
                    VerticalAlignment="Center"/>
        </StackPanel>
      </Grid>

      <Border Grid.Row="1" Margin="0,0,0,8" Background="{StaticResource MallasPanelBg}"
              BorderBrush="{StaticResource MallasBorder}"
              BorderThickness="1" CornerRadius="4" Padding="8,6">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock x:Name="TxtMallasSummary" Grid.Row="0"
                     Foreground="{StaticResource MallasAccent}" FontSize="10" FontWeight="SemiBold"/>
          <TextBlock x:Name="TxtMallasDetail" Grid.Row="1" Margin="0,4,0,0"
                     Foreground="{StaticResource MallasFgLo}" FontSize="10" TextWrapping="Wrap"/>
        </Grid>
      </Border>

      <TextBlock x:Name="TxtInfoMuros" Grid.Row="2"
                 Foreground="{StaticResource MallasFgLo}" FontSize="10"
                 TextWrapping="Wrap" Margin="0,0,0,10"
                 Text="Clic fuste · Ctrl+clic multi · Mayús+clic rango · arrastre en el alzado = marquee."/>

      <StackPanel x:Name="PnlMallasHiddenCtrls" Grid.Row="3" Visibility="Collapsed"/>

      <Grid Grid.Row="3">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="__SECTION_RAIL_WIDTH__"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="{StaticResource MallasPanelBg}"
                BorderBrush="{StaticResource MallasBorder}" BorderThickness="1"
                CornerRadius="4,0,0,4" Padding="0">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border x:Name="BdrElevCotaToolbar" Grid.Row="0" Visibility="Collapsed"
                    Background="#0E1B32" BorderBrush="{StaticResource MallasBorder}"
                    BorderThickness="0,0,0,1" Padding="8,5,8,5"
                    ToolTip="Referencia de cota en el alzado">
              <StackPanel x:Name="PnlElevCotaToolbar" Orientation="Horizontal"
                          VerticalAlignment="Center"/>
            </Border>
            <Border x:Name="BdrColumnHeaders" Grid.Row="1"
                    Background="{StaticResource MallasPanelBg}"
                    BorderBrush="{StaticResource MallasBorder}"
                    BorderThickness="0,0,0,1" Padding="8,6,8,4">
              <Grid x:Name="GrdColumnHeaders" Background="Transparent"
                    SnapsToDevicePixels="True" HorizontalAlignment="Center"/>
            </Border>
            <ScrollViewer x:Name="ScrMuros" Grid.Row="2"
                          VerticalScrollBarVisibility="Auto"
                          HorizontalScrollBarVisibility="Disabled"
                          CanContentScroll="False"
                          IsDeferredScrollingEnabled="True">
              <Border Background="{StaticResource MallasPanelBg}"
                      BorderBrush="Transparent" BorderThickness="0"
                      Padding="8,4,8,12">
                <Grid x:Name="GrdListaMuros" Background="Transparent"
                      ClipToBounds="False" SnapsToDevicePixels="True"
                      HorizontalAlignment="Center"/>
              </Border>
            </ScrollViewer>
          </Grid>
        </Border>

        <Border Grid.Column="1" Background="{StaticResource MallasPanelBg}"
                BorderBrush="{StaticResource MallasBorder}" BorderThickness="1,1,1,1"
                CornerRadius="0,4,4,0" Padding="8,8">
          <ScrollViewer VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Disabled"
                        CanContentScroll="False"
                        IsDeferredScrollingEnabled="True">
            <Grid x:Name="PnlSectionRail" SnapsToDevicePixels="True">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Grid x:Name="GrdSectionRailHint" Grid.Row="0" Margin="0,0,0,8"
                    Visibility="Collapsed">
                <TextBlock x:Name="TxtSectionRailHint"
                           Text="Configuración · vertical y horizontal"
                           Foreground="{StaticResource MallasFgLo}" FontSize="9"
                           VerticalAlignment="Center" TextWrapping="Wrap"/>
              </Grid>
              <StackPanel x:Name="GrdRailHeader" Grid.Row="0" Visibility="Collapsed">
                <CheckBox x:Name="ChkPhaseMallas" Visibility="Collapsed"/>
              </StackPanel>
              <StackPanel x:Name="PnlSectionCtrls" Grid.Row="1" Margin="0,0,0,0"
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
                   Foreground="{StaticResource MallasFgLo}" FontSize="10"
                   TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnManual" Content="Manual"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="100"
                  Background="#2A5C3D" Margin="0,0,10,0"
                  ToolTip="Abrir manual de usuario"/>
          <Button x:Name="BtnCancelar" Content="Cancelar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnCrear" Content="Crear mallas"
                  Style="{StaticResource BtnPrimary}" MinWidth="180"
                  ToolTip="Crea Area Reinforcement según toggles y parámetros del rail"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>"""


def build_mallas_en_muros_xaml(styles_xml=None):
    xaml = XAML_MALLAS_EN_MUROS
    xaml = xaml.replace(u"__BIMTOOLS_DARK_STYLES__", styles_xml or BIMTOOLS_DARK_STYLES_XML)
    xaml = xaml.replace(u"__WINDOW_CHROME_TITLE__", WINDOW_CHROME_TITLE)
    xaml = xaml.replace(
        u"__SECTION_RAIL_WIDTH__",
        u"{0:.0f}".format(float(MALLAS_SECTION_RAIL_WIDTH_PX)),
    )
    xaml = xaml.replace(u"__BG_APP__", BG_APP)
    xaml = xaml.replace(u"__FONT_FAMILY__", FONT_FAMILY)
    xaml = xaml.replace(u"__FONT_SIZE_BASE__", str(FONT_SIZE_BASE))
    return xaml
