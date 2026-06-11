# -*- coding: utf-8 -*-
"""
Ventana modal para gestionar personas.json (desde Revisiones).
"""

from __future__ import print_function

import json
import os

try:
    unicode
except NameError:
    unicode = str

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.IO")

from System.Collections import IComparer
from System.IO import Directory
from System.Text import Encoding
from System.Windows.Markup import XamlReader
from System.Windows import RoutedEventHandler

import System.IO as sio

from infra.revit_wpf_window_position import (  # noqa: E402
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

try:
    from infra.bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML  # noqa: E402
except ImportError:
    BIMTOOLS_DARK_STYLES_XML = u""

PERSONA_ROL_MODELADOR = u"Modelador"
PERSONA_ROL_INGENIERO = u"Ingeniero"

# Mismo aspecto que 03_Personas.pushbutton (compatible con bimtools_wpf_dark_theme).
PERSONAS_REVISIONES_LIKE_STYLES = u"""
    <Style x:Key="PanelInset" TargetType="Border">
      <Setter Property="Background" Value="#071018"/>
      <Setter Property="BorderBrush" Value="#1E3F55"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="12,10"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
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
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource {x:Type DataGridColumnHeader}}">
      <Setter Property="Background" Value="#0F2840"/>
      <Setter Property="Foreground" Value="#C8E4EF"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Padding" Value="12,10"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
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
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
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
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="MinHeight" Value="32"/>
      <Setter Property="CaretBrush" Value="#7AA3B8"/>
    </Style>
    <Style x:Key="PersonaFormCombo" TargetType="ComboBox" BasedOn="{StaticResource Combo}">
      <Setter Property="Width" Value="{x:Static sys:Double.NaN}"/>
      <Setter Property="Height" Value="{x:Static sys:Double.NaN}"/>
      <Setter Property="HorizontalAlignment" Value="Stretch"/>
      <Setter Property="MinWidth" Value="0"/>
      <Setter Property="MaxWidth" Value="99999"/>
      <Setter Property="MinHeight" Value="32"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="ItemContainerStyle" Value="{StaticResource ComboItem}"/>
    </Style>
    <Style x:Key="PersonaGridRolCombo" TargetType="ComboBox" BasedOn="{StaticResource PersonaFormCombo}">
      <Setter Property="HorizontalAlignment" Value="Center"/>
      <Setter Property="Width" Value="118"/>
      <Setter Property="MinWidth" Value="100"/>
      <Setter Property="MaxWidth" Value="128"/>
    </Style>
    <Style x:Key="PersonaGridCellTextBlock" TargetType="TextBlock">
      <Setter Property="TextAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="PersonaGridEditTextBox" TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
      <Setter Property="TextAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="ExpLamScrollBarDark" TargetType="ScrollBar" BasedOn="{StaticResource BimToolsScrollBarDark}"/>
"""

GESTIONAR_PERSONAS_XAML = (
    u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:sys="clr-namespace:System;assembly=mscorlib"
    x:Name="PersonasWin"
    Title="Personas"
    Height="620" Width="680" MinHeight="500" MinWidth="560"
    Background="Transparent"
    AllowsTransparency="True"
    WindowStyle="None"
    ResizeMode="CanResize"
    WindowStartupLocation="Manual"
    FontFamily="Segoe UI" FontSize="12" UseLayoutRounding="True">
  <Window.Resources>
"""
    + BIMTOOLS_DARK_STYLES_XML
    + PERSONAS_REVISIONES_LIKE_STYLES
    + u"""
    <x:Array x:Key="PersonaRolOpciones" Type="sys:String" xmlns:sys="clr-namespace:System;assembly=mscorlib">
      <sys:String>Modelador</sys:String>
      <sys:String>Ingeniero</sys:String>
    </x:Array>
  </Window.Resources>
  <Border x:Name="PersonasRootChrome" CornerRadius="8" Background="#0E1B32" Padding="14"
          BorderBrush="#5BC0DE" BorderThickness="1" ClipToBounds="True">
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
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Image x:Name="ImgLogo" Width="40" Height="40" Grid.Column="0"
                 Stretch="Uniform" Margin="0,0,10,0" VerticalAlignment="Center"/>
          <StackPanel Grid.Column="1" VerticalAlignment="Center">
            <TextBlock Text="Personas" FontSize="14" FontWeight="Bold" Foreground="#FFFFFF"/>
            <TextBlock Text="Directorio compartido · personas.json" FontSize="11" Foreground="#95B8CC" Margin="0,4,0,0"/>
          </StackPanel>
          <Button x:Name="BtnClose" Grid.Column="2"
                  Style="{StaticResource BtnCloseX_MinimalNoBg}"
                  VerticalAlignment="Center" HorizontalAlignment="Right" ToolTip="Cerrar"/>
        </Grid>
      </Border>
      <TextBlock Grid.Row="1" Text="Doble clic en una celda para editar; los cambios se guardan al salir de la celda."
                 Foreground="#6B94AA" FontSize="10" Margin="0,0,0,10" TextWrapping="Wrap"/>
      <Border Grid.Row="2" Style="{StaticResource PanelInset}" Padding="14,12">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="16"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,0,10">
            <TextBlock Text="Nombre *" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <TextBox x:Name="TxtNombreNew"/>
          </StackPanel>
          <StackPanel Grid.Row="0" Grid.Column="2" Margin="0,0,0,10">
            <TextBlock Text="Email" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <TextBox x:Name="TxtEmailNew"/>
          </StackPanel>
          <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,0,0,12">
            <TextBlock Text="Abreviación (ej. J.N.N.)" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <TextBox x:Name="TxtAbreviacionNew" MaxLength="32"/>
          </StackPanel>
          <StackPanel Grid.Row="1" Grid.Column="2" Margin="0,0,0,12">
            <TextBlock Text="Rol *" Style="{StaticResource LabelSmall}" Margin="0,0,0,4"/>
            <ComboBox x:Name="CmbRolNew" Style="{StaticResource PersonaFormCombo}" IsEditable="False"/>
          </StackPanel>
          <StackPanel Grid.Row="2" Grid.Column="2" Margin="0,4,0,0">
            <Button x:Name="BtnAgregar" Content="+ Agregar al directorio"
                    Style="{StaticResource BtnPrimary}" MinWidth="168" Padding="16,8" HorizontalAlignment="Right"/>
          </StackPanel>
        </Grid>
      </Border>
      <Border Grid.Row="3" Background="#040A12" BorderBrush="#1E3F55" BorderThickness="1" CornerRadius="8" Padding="0" Margin="0,0,0,10">
        <DataGrid x:Name="GridPersonas"
                  AutoGenerateColumns="False" IsReadOnly="False"
                  CanUserSortColumns="True"
                  SelectionMode="Single" CanUserAddRows="False"
                  AlternationCount="2" RowHeaderWidth="0">
          <DataGrid.CellStyle>
            <Style TargetType="DataGridCell" BasedOn="{StaticResource GridCellPadding}"/>
          </DataGrid.CellStyle>
          <DataGrid.Columns>
            <DataGridTextColumn Header="Nombre" Width="2*" MinWidth="120" SortMemberPath="Nombre"
                ElementStyle="{StaticResource PersonaGridCellTextBlock}"
                EditingElementStyle="{StaticResource PersonaGridEditTextBox}">
              <DataGridTextColumn.Binding>
                <Binding Path="Nombre" Mode="TwoWay" UpdateSourceTrigger="LostFocus"/>
              </DataGridTextColumn.Binding>
            </DataGridTextColumn>
            <DataGridTextColumn Header="Abrev." Width="80" MinWidth="72" MaxWidth="100" SortMemberPath="Abreviacion"
                ElementStyle="{StaticResource PersonaGridCellTextBlock}"
                EditingElementStyle="{StaticResource PersonaGridEditTextBox}">
              <DataGridTextColumn.Binding>
                <Binding Path="Abreviacion" Mode="TwoWay" UpdateSourceTrigger="LostFocus"/>
              </DataGridTextColumn.Binding>
            </DataGridTextColumn>
            <DataGridComboBoxColumn Header="Rol" Width="128" MinWidth="120" SortMemberPath="Rol"
                SelectedItemBinding="{Binding Rol, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                ItemsSource="{StaticResource PersonaRolOpciones}">
              <DataGridComboBoxColumn.ElementStyle>
                <Style TargetType="ComboBox" BasedOn="{StaticResource PersonaGridRolCombo}">
                  <Setter Property="IsHitTestVisible" Value="False"/>
                  <Setter Property="Focusable" Value="False"/>
                  <Setter Property="IsTabStop" Value="False"/>
                </Style>
              </DataGridComboBoxColumn.ElementStyle>
              <DataGridComboBoxColumn.EditingElementStyle>
                <Style TargetType="ComboBox" BasedOn="{StaticResource PersonaGridRolCombo}"/>
              </DataGridComboBoxColumn.EditingElementStyle>
            </DataGridComboBoxColumn>
            <DataGridTextColumn Header="Email" Width="3*" MinWidth="200" SortMemberPath="Email"
                ElementStyle="{StaticResource PersonaGridCellTextBlock}"
                EditingElementStyle="{StaticResource PersonaGridEditTextBox}">
              <DataGridTextColumn.Binding>
                <Binding Path="Email" Mode="TwoWay" UpdateSourceTrigger="LostFocus"/>
              </DataGridTextColumn.Binding>
            </DataGridTextColumn>
          </DataGrid.Columns>
        </DataGrid>
      </Border>
      <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,4,0,0">
        <Button x:Name="BtnEliminar" Content="Eliminar seleccionado"
                Style="{StaticResource BtnGhost}" Margin="0,0,10,0"/>
        <Button x:Name="BtnCerrar" Content="Cerrar"
                Style="{StaticResource BtnPrimary}" Padding="18,9"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
"""
)


def _normalize_rol(val):
    s = (val or u"").strip()
    if s == PERSONA_ROL_INGENIERO:
        return PERSONA_ROL_INGENIERO
    return PERSONA_ROL_MODELADOR


class PersonaItem(object):
    def __init__(self, nombre, email=u"", abreviacion=u"", rol=u""):
        self.Nombre = nombre
        self.Email = email
        self.Abreviacion = abreviacion or u""
        self.Rol = _normalize_rol(rol)

    def __str__(self):
        return self.Nombre

    def ToString(self):
        return self.Nombre


class _PersonaGridComparer(IComparer):
    def __init__(self, prop_name, ascending):
        self._prop = prop_name
        self._mul = 1 if ascending else -1

    def _key(self, p):
        try:
            v = getattr(p, self._prop, None)
            s = u"" if v is None else str(v)
            return s.strip().lower()
        except Exception:
            return u""

    def Compare(self, x, y):
        a, b = self._key(x), self._key(y)
        if a < b:
            return -1 * self._mul
        if a > b:
            return 1 * self._mul
        return 0


def _wire_personas_data_grid_sort(grid, personas_coll):
    def on_sort(sender, e):
        col = e.Column
        if col is None:
            return
        path = col.SortMemberPath
        if not path:
            return
        try:
            path = str(path)
        except Exception:
            return

        from System.ComponentModel import ListSortDirection
        from System.Windows.Data import CollectionViewSource, ListCollectionView

        sd = col.SortDirection
        if sd is None:
            ascending = True
        else:
            ascending = sd == ListSortDirection.Ascending

        e.Handled = True

        view = CollectionViewSource.GetDefaultView(personas_coll)
        try:
            if isinstance(view, ListCollectionView):
                view.CustomSort = _PersonaGridComparer(path, ascending)
            else:
                items = list(personas_coll)
                items.sort(
                    key=lambda p: (getattr(p, path, None) or u"").strip().lower(),
                    reverse=not ascending,
                )
                personas_coll.Clear()
                for it in items:
                    personas_coll.Add(it)
        except Exception:
            pass

        try:
            for c in sender.Columns:
                if c is not col:
                    c.SortDirection = None
        except Exception:
            pass

    grid.Sorting += on_sort


def load_personas_list(personas_file):
    personas = []
    if os.path.isfile(personas_file):
        try:
            raw = sio.File.ReadAllText(personas_file, Encoding.UTF8)
            data = json.loads(raw)
            for p in data:
                personas.append(
                    PersonaItem(
                        p.get("nombre", ""),
                        p.get("email", ""),
                        p.get("abreviacion", ""),
                        p.get("rol", ""),
                    )
                )
        except Exception:
            pass
    return personas


def save_personas_list(personas_list, issues_dir, personas_file):
    try:
        Directory.CreateDirectory(issues_dir)
    except Exception:
        pass
    data = [
        {
            "nombre": p.Nombre,
            "email": p.Email,
            "abreviacion": getattr(p, "Abreviacion", "") or "",
            "rol": _normalize_rol(getattr(p, "Rol", "")),
        }
        for p in personas_list
    ]
    sio.File.WriteAllText(
        personas_file,
        json.dumps(data, ensure_ascii=False, indent=2),
        Encoding.UTF8,
    )


def _personas_apply_title_chrome(win):
    try:
        from System.Windows.Input import MouseButtonEventHandler

        btn_close = win.FindName("BtnClose")
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
        import infra.bimtools_paths as bimtools_paths

        img = win.FindName("ImgLogo")
        if img is not None:
            bmp = bimtools_paths.load_logo_bitmap_image()
            if bmp is not None:
                img.Source = bmp
    except Exception:
        pass


class GestionarPersonasDialog(object):
    """Modal: alta, edición en grilla y baja; persiste en personas.json."""

    def __init__(
        self,
        personas_collection,
        issues_dir,
        personas_file,
        uidoc=None,
        revit_app=None,
        owner=None,
    ):
        self._personas = personas_collection
        self._issues_dir = issues_dir
        self._personas_file = personas_file
        win = XamlReader.Parse(GESTIONAR_PERSONAS_XAML)
        self._win = win

        self._txt_nombre = win.FindName("TxtNombreNew")
        self._txt_email = win.FindName("TxtEmailNew")
        self._txt_abrev = win.FindName("TxtAbreviacionNew")
        self._cmb_rol = win.FindName("CmbRolNew")
        if self._cmb_rol is not None:
            self._cmb_rol.Items.Clear()
            self._cmb_rol.Items.Add(PERSONA_ROL_MODELADOR)
            self._cmb_rol.Items.Add(PERSONA_ROL_INGENIERO)
            self._cmb_rol.SelectedIndex = 0
        self._btn_add = win.FindName("BtnAgregar")
        self._grid = win.FindName("GridPersonas")
        self._btn_del = win.FindName("BtnEliminar")
        self._btn_close = win.FindName("BtnCerrar")
        self._btn_title_close = win.FindName("BtnClose")

        self._grid.ItemsSource = self._personas
        self._grid_edit_backup = None
        self._grid.BeginningEdit += self._on_grid_beginning_edit
        self._grid.CellEditEnding += self._on_grid_cell_edit_ending
        _wire_personas_data_grid_sort(self._grid, self._personas)

        self._btn_add.Click += RoutedEventHandler(self._on_agregar)
        self._btn_del.Click += RoutedEventHandler(self._on_eliminar)
        self._btn_close.Click += RoutedEventHandler(self._on_cerrar)
        if self._btn_title_close is not None:
            self._btn_title_close.Click += RoutedEventHandler(self._on_cerrar)

        _personas_apply_title_chrome(win)

        from System.Windows import WindowStartupLocation

        has_owner = False
        if owner is not None:
            try:
                win.Owner = owner
                has_owner = True
            except Exception:
                pass
        try:
            if has_owner:
                win.WindowStartupLocation = WindowStartupLocation.CenterOwner
            else:
                win.WindowStartupLocation = WindowStartupLocation.Manual
                if uidoc is not None and revit_app is not None:
                    _hw = revit_main_hwnd(revit_app)
                    position_wpf_window_top_left_at_active_view(win, uidoc, _hw)
        except Exception:
            try:
                win.WindowStartupLocation = WindowStartupLocation.CenterScreen
            except Exception:
                pass
        try:
            win.ShowActivated = True
        except Exception:
            pass
        win.ShowDialog()

    def _on_grid_beginning_edit(self, sender, e):
        row = e.Row.Item
        if isinstance(row, PersonaItem):
            self._grid_edit_backup = (row.Nombre, row.Email, row.Abreviacion, row.Rol)

    def _on_grid_cell_edit_ending(self, sender, e):
        try:
            if int(e.EditAction) != 1:
                return
        except Exception:
            return
        row = e.Row.Item
        if not isinstance(row, PersonaItem):
            return
        row.Nombre = (row.Nombre or "").strip()
        row.Email = (row.Email or "").strip()
        row.Abreviacion = (row.Abreviacion or "").strip()
        row.Rol = _normalize_rol(row.Rol)
        if not row.Nombre:
            b = self._grid_edit_backup
            if b:
                row.Nombre = b[0]
                row.Email = b[1]
                row.Abreviacion = b[2]
                row.Rol = b[3]
            try:
                e.Cancel = True
            except Exception:
                pass
            return
        save_personas_list(
            list(self._personas), self._issues_dir, self._personas_file
        )

    def _on_agregar(self, sender, args):
        nombre = self._txt_nombre.Text.strip()
        if not nombre:
            return
        email = self._txt_email.Text.strip()
        abrev = self._txt_abrev.Text.strip() if self._txt_abrev else ""
        rol_sel = PERSONA_ROL_MODELADOR
        if self._cmb_rol is not None and self._cmb_rol.SelectedItem is not None:
            try:
                rol_sel = _normalize_rol(str(self._cmb_rol.SelectedItem))
            except Exception:
                rol_sel = PERSONA_ROL_MODELADOR
        self._personas.Add(PersonaItem(nombre, email, abrev, rol_sel))
        save_personas_list(
            list(self._personas), self._issues_dir, self._personas_file
        )
        self._txt_nombre.Text = ""
        self._txt_email.Text = ""
        if self._txt_abrev:
            self._txt_abrev.Text = ""
        if self._cmb_rol is not None:
            self._cmb_rol.SelectedIndex = 0
        self._txt_nombre.Focus()

    def _on_eliminar(self, sender, args):
        sel = self._grid.SelectedItem
        if sel:
            self._personas.Remove(sel)
            save_personas_list(
                list(self._personas), self._issues_dir, self._personas_file
            )

    def _on_cerrar(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass
