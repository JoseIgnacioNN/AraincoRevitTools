# -*- coding: utf-8 -*-
"""
BIMIssue — Crea incidencias con screenshot, viewpoint y lista de elementos.
Inspirado en BIMcollab. Guarda cada issue en una carpeta con PNG + JSON.
"""

__title__ = "Crear\nIncidencia"
__author__ = "pyRevit"
__doc__    = "Crea issues con screenshot de la vista activa, elementos seleccionados y metadatos."

import os
import json
import datetime
import tempfile

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System")
clr.AddReference("System.Collections")
clr.AddReference("System.Runtime.InteropServices")

from Autodesk.Revit.DB import ElementId, View3D, ModelPathUtils
from System.Drawing         import Bitmap, Graphics, Rectangle, Size as DSize
from System.Drawing.Imaging import ImageFormat
from System.Windows.Forms   import Screen
import System.Drawing
from Autodesk.Revit.UI.Selection import ObjectType

from System.Windows.Markup   import XamlReader
from System.Windows          import Window, MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, RoutedEventHandler
from System.Windows.Input   import MouseButtonEventHandler, Key, KeyEventHandler
from System.Windows.Controls import TextSearch, SelectionChangedEventHandler
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Media.Imaging   import BitmapImage
from System                  import Uri, UriKind
import System.IO              as sio
import System

doc   = __revit__.ActiveUIDocument.Document   # noqa
uidoc = __revit__.ActiveUIDocument            # noqa

ISSUES_DIR    = u"Y:\\00_SERVIDOR DE INCIDENCIAS"
PERSONAS_FILE = os.path.join(ISSUES_DIR, "personas.json")

# Paleta Arainco (igual que Recordatorios)
_COLOR_DARK = u"#264A62"   # Azul oscuro - titulos
_COLOR_LIGHT = u"#51B2E0"  # Azul claro - acentos
_COLOR_GRAY = u"#B4B4B4"   # Gris - bordes
_COLOR_WHITE = u"#FFFFFF"


def _get_project_folder_name(document):
    """
    Devuelve un nombre de carpeta limpio para el proyecto activo.

    Formato requerido:
        "DIRECTORIO QUE CONTIENE EL MODELO CENTRAL" _ "NOMBRE DEL MODELO CENTRAL"

    Lógica:
    1. Del archivo actual se obtiene el modelo central (si es workshared) o se usa
       la ruta del archivo actual (si es independiente).
    2. Del modelo central se obtiene: el nombre del directorio que lo contiene y
       el nombre del archivo del modelo central (sin extensión).
    Los caracteres inválidos para nombres de carpeta se reemplazan por '_'.
    """
    import os
    import re
    folder_name = u""
    try:
        model_title = None
        try:
            model_title = (document.Title or u"").strip()
        except Exception:
            model_title = None

        def _from_path(path, title_fallback):
            if not path:
                base = (title_fallback or u"SinNombre").strip() or u"SinNombre"
                return u"{}_{}".format(base, base)
            # Directorio que contiene el modelo central
            dir_containing = os.path.basename(os.path.dirname(path))
            if not dir_containing:
                dir_containing = (title_fallback or u"SinNombre").strip() or u"SinNombre"
            # Nombre del modelo central (archivo sin extensión)
            model_name = os.path.splitext(os.path.basename(path))[0] or u"SinNombre"
            return u"{}_{}".format(dir_containing, model_name)

        if document.IsWorkshared:
            central_path = document.GetWorksharingCentralModelPath()
            user_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(central_path)
            folder_name = _from_path(user_path, model_title)
        if not folder_name and document.PathName:
            # Archivo independiente: la ruta del archivo actual es la del "modelo central"
            folder_name = _from_path(document.PathName, model_title)
        if not folder_name:
            base = (model_title or u"SinNombre").strip() or u"SinNombre"
            folder_name = u"{0}_{0}".format(base)
    except Exception:
        try:
            base = (document.Title or u"SinNombre").strip() or u"SinNombre"
            folder_name = u"{0}_{0}".format(base)
        except Exception:
            folder_name = u"SinNombre_SinNombre"
    # Eliminar caracteres inválidos en nombres de carpeta Windows
    folder_name = re.sub(r'[<>:"/\\|?*]', u"_", folder_name).strip(u". ")
    return folder_name if folder_name else u"SinNombre_SinNombre"


# ── Modelo de persona ────────────────────────────────────────────────────────
class PersonaItem(object):
    def __init__(self, nombre, email=""):
        self.Nombre = nombre
        self.Email  = email

    def __str__(self):
        return self.Nombre

    def ToString(self):
        return self.Nombre


def _load_personas(personas_file):
    """Lee personas.json y devuelve lista de PersonaItem."""
    import os
    import json
    personas = []
    if os.path.exists(personas_file):
        try:
            raw  = sio.File.ReadAllText(personas_file, System.Text.Encoding.UTF8)
            data = json.loads(raw)
            for p in data:
                personas.append(PersonaItem(p.get("nombre", ""), p.get("email", "")))
        except Exception:
            pass
    return personas


def _save_personas(personas, issues_dir, personas_file):
    """Serializa la lista de PersonaItem a personas.json."""
    import json
    sio.Directory.CreateDirectory(issues_dir)
    data = [{"nombre": p.Nombre, "email": p.Email} for p in personas]
    sio.File.WriteAllText(
        personas_file,
        json.dumps(data, ensure_ascii=False, indent=2),
        System.Text.Encoding.UTF8,
    )


# ── XAML — diálogo Gestionar Personas ───────────────────────────────────────
GESTIONAR_PERSONAS_XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Personas"
    Height="520" Width="540"
    WindowStartupLocation="CenterOwner"
    Background="#0A1C26"
    FontFamily="Segoe UI"
    ResizeMode="CanResize">
  <Grid Margin="20,16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Text="DIRECTORIO DE PERSONAS"
               Foreground="#4A8BA6" FontSize="11" FontWeight="SemiBold"
               Margin="0,0,0,14"/>

    <!-- Formulario agregar -->
    <Border Grid.Row="1" Background="#0F2535" CornerRadius="6"
            Padding="14,12" Margin="0,0,0,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Margin="0,0,8,0">
          <TextBlock Text="Nombre *" Foreground="#4A8BA6" FontSize="10"
                     FontWeight="SemiBold" Margin="0,0,0,4"/>
          <TextBox x:Name="TxtNombreNew"
                   Background="#0A1C26" Foreground="#C8E4EF"
                   CaretBrush="#5BB8D4"
                   BorderBrush="#1A3D52" BorderThickness="1"
                   Padding="8,7" FontSize="12"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Margin="0,0,10,0">
          <TextBlock Text="Email" Foreground="#4A8BA6" FontSize="10"
                     FontWeight="SemiBold" Margin="0,0,0,4"/>
          <TextBox x:Name="TxtEmailNew"
                   Background="#0A1C26" Foreground="#C8E4EF"
                   CaretBrush="#5BB8D4"
                   BorderBrush="#1A3D52" BorderThickness="1"
                   Padding="8,7" FontSize="12"/>
        </StackPanel>
        <Button x:Name="BtnAgregar" Grid.Column="2" Content="+ Agregar"
                VerticalAlignment="Bottom" Padding="14,8"
                Background="#5BB8D4" Foreground="#0A1C26"
                FontWeight="Bold" FontSize="12"
                BorderThickness="0" Cursor="Hand"/>
      </Grid>
    </Border>

    <!-- Lista -->
    <DataGrid Grid.Row="2" x:Name="GridPersonas"
              AutoGenerateColumns="False" IsReadOnly="True"
              SelectionMode="Single" CanUserAddRows="False"
              Background="#081520" Foreground="#C8E4EF"
              BorderBrush="#1A3D52" BorderThickness="1"
              GridLinesVisibility="Horizontal"
              HorizontalGridLinesBrush="#1A3D52"
              RowBackground="#081520"
              AlternatingRowBackground="#0F2535"
              ColumnHeaderHeight="30" RowHeight="34" FontSize="12">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Nombre" Binding="{Binding Nombre}" Width="*"/>
        <DataGridTextColumn Header="Email"  Binding="{Binding Email}"  Width="*"/>
      </DataGrid.Columns>
      <DataGrid.ColumnHeaderStyle>
        <Style TargetType="DataGridColumnHeader">
          <Setter Property="Background"  Value="#1A3D52"/>
          <Setter Property="Foreground"  Value="#5BB8D4"/>
          <Setter Property="FontWeight"  Value="SemiBold"/>
          <Setter Property="Padding"     Value="10,0"/>
          <Setter Property="BorderBrush" Value="#2A5570"/>
          <Setter Property="BorderThickness" Value="0,0,1,0"/>
        </Style>
      </DataGrid.ColumnHeaderStyle>
      <DataGrid.CellStyle>
        <Style TargetType="DataGridCell">
          <Setter Property="BorderThickness" Value="0"/>
          <Setter Property="Padding"         Value="10,0"/>
          <Style.Triggers>
            <Trigger Property="IsSelected" Value="True">
              <Setter Property="Background" Value="#1E4D66"/>
            </Trigger>
          </Style.Triggers>
        </Style>
      </DataGrid.CellStyle>
    </DataGrid>

    <!-- Pie -->
    <StackPanel Grid.Row="3" Orientation="Horizontal"
                HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button x:Name="BtnEliminar" Content="Eliminar seleccionado"
              Padding="14,8" Margin="0,0,10,0"
              Background="#1A3D52" Foreground="#C8E4EF"
              BorderThickness="0" FontSize="12" Cursor="Hand"/>
      <Button x:Name="BtnCerrar" Content="Cerrar"
              Padding="20,8"
              Background="#1A4F6A" Foreground="#C8E4EF"
              FontWeight="Bold" BorderThickness="0"
              FontSize="12" Cursor="Hand"/>
    </StackPanel>
  </Grid>
</Window>
"""


# ── XAML — ventana principal ─────────────────────────────────────────────────
XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Crear Incidencia"
    Height="1150" Width="1020"
    MinHeight="1150" MinWidth="900"
    WindowStartupLocation="CenterScreen"
    Background="#0A1C26"
    FontFamily="Segoe UI"
    ResizeMode="CanResize">

  <Window.Resources>
    <!--
        Paleta derivada del logo corporativo:
          Azul celeste  #5BB8D4   ← trazo principal del logo
          Teal oscuro   #1A4F6A   ← triángulos derechos del logo
          Gris cálido   #B8AFA8   ← barra horizontal del logo
    -->
    <Style x:Key="Label" TargetType="TextBlock">
      <Setter Property="Foreground"  Value="#4A8BA6"/>
      <Setter Property="FontSize"    Value="11"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Margin"      Value="0,10,0,4"/>
    </Style>
    <Style x:Key="Input" TargetType="TextBox">
      <Setter Property="Background"      Value="#0F2535"/>
      <Setter Property="Foreground"      Value="#C8E4EF"/>
      <Setter Property="BorderBrush"     Value="#1A3D52"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="10,6"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="CaretBrush"      Value="#5BB8D4"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="6" Padding="{TemplateBinding Padding}">
              <Grid>
                <TextBlock x:Name="Watermark"
                           Text="{TemplateBinding Tag}"
                           Foreground="#2A5C75" FontStyle="Italic"
                           IsHitTestVisible="False" Visibility="Collapsed"
                           VerticalAlignment="Center"/>
                <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="Center"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <DataTrigger Binding="{Binding Text, RelativeSource={RelativeSource Self}}" Value="">
                <Setter TargetName="Watermark" Property="Visibility" Value="Visible"/>
              </DataTrigger>
              <Trigger Property="IsFocused" Value="True">
                <Setter TargetName="Watermark" Property="Visibility" Value="Collapsed"/>
                <Setter Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <!-- ComboBox no editable (Prioridad, Estado) -->
    <Style x:Key="Combo" TargetType="ComboBox">
      <Setter Property="Background"      Value="#0D2234"/>
      <Setter Property="Foreground"      Value="#F0F8FC"/>
      <Setter Property="BorderBrush"     Value="#1A3D52"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Height"          Value="36"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid TextElement.Foreground="{TemplateBinding Foreground}">
              <Border x:Name="Border"
                      Background="{TemplateBinding Background}"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}"
                      CornerRadius="6"/>
              <ToggleButton
                  IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                  Focusable="False">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border Background="Transparent"/>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter x:Name="ContentSite"
                                IsHitTestVisible="False"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                Margin="12,0,32,0"
                                VerticalAlignment="Center"/>
              <TextBlock Text="&#9660;" FontSize="9" Foreground="#5BB8D4"
                         HorizontalAlignment="Right" VerticalAlignment="Center"
                         Margin="0,0,12,0" IsHitTestVisible="False"/>
              <Popup x:Name="PART_Popup"
                     IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False"
                     PopupAnimation="Fade" Placement="Bottom">
                <Border Background="#0D2234" BorderBrush="#1A3D52" BorderThickness="1"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <ScrollViewer MaxHeight="200" VerticalScrollBarVisibility="Auto">
                    <ItemsPresenter/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Border" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="Border" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ComboBox solo selección (Reportado por, Asignado a) -->
    <Style x:Key="ComboPersona" TargetType="ComboBox">
      <Setter Property="Background"      Value="#0D2234"/>
      <Setter Property="Foreground"      Value="#F0F8FC"/>
      <Setter Property="BorderBrush"     Value="#1A3D52"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Height"          Value="36"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <Border x:Name="Border"
                      Background="{TemplateBinding Background}"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}"
                      CornerRadius="6"/>
              <Grid Margin="12,0,32,0">
                <TextBox x:Name="PART_EditableTextBox"
                         Background="Transparent"
                         Foreground="#F0F8FC"
                         CaretBrush="#5BB8D4"
                         BorderThickness="0"
                         VerticalAlignment="Center"
                         IsReadOnly="True"
                         Focusable="False"/>
              </Grid>
              <ToggleButton HorizontalAlignment="Right" Width="30" Cursor="Hand"
                            IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                            Focusable="False">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border Background="Transparent">
                      <TextBlock Text="&#9660;" FontSize="9" Foreground="#5BB8D4"
                                 HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <Popup x:Name="PART_Popup"
                     IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False"
                     PopupAnimation="Fade" Placement="Bottom">
                <Border Background="#0D2234" BorderBrush="#1A3D52" BorderThickness="1"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <ScrollViewer MaxHeight="240" VerticalScrollBarVisibility="Auto">
                    <ItemsPresenter/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Border" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="Border" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Background"      Value="#5BB8D4"/>
      <Setter Property="Foreground"      Value="#0A1C26"/>
      <Setter Property="FontWeight"      Value="Bold"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Padding"         Value="20,9"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="7"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#7CCDE2"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="#4AA5C0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnGhost" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="Background" Value="#0F2535"/>
      <Setter Property="Foreground" Value="#C8E4EF"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#1A3D52"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="BtnSuccess" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="Background" Value="#1A4F6A"/>
      <Setter Property="Foreground" Value="#C8E4EF"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#235E7D"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="BtnLink" TargetType="Button">
      <Setter Property="Background"      Value="Transparent"/>
      <Setter Property="Foreground"      Value="#4A8BA6"/>
      <Setter Property="FontSize"        Value="11"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Padding"         Value="0,3"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <TextBlock x:Name="Lbl"
                       Text="{TemplateBinding Content}"
                       Foreground="{TemplateBinding Foreground}"
                       FontSize="{TemplateBinding FontSize}"
                       Padding="{TemplateBinding Padding}"
                       TextDecorations="Underline"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Lbl" Property="Foreground" Value="#5BB8D4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="ChkDisciplina" TargetType="CheckBox">
      <Setter Property="Foreground"  Value="#C8E4EF"/>
      <Setter Property="FontSize"    Value="12"/>
      <Setter Property="Margin"      Value="0,0,12,6"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Border x:Name="Bd" Background="#1A3D52" CornerRadius="5"
                    Padding="10,5" BorderThickness="1" BorderBrush="#2A5570">
              <StackPanel Orientation="Horizontal">
                <Border x:Name="Check" Width="14" Height="14" CornerRadius="3"
                        Background="Transparent" BorderBrush="#4A8BA6"
                        BorderThickness="1.5" Margin="0,0,7,0" VerticalAlignment="Center">
                  <TextBlock x:Name="CheckMark" Text="&#10003;" FontSize="10"
                             FontWeight="Bold" Foreground="#0A1C26"
                             HorizontalAlignment="Center" VerticalAlignment="Center"
                             Visibility="Collapsed"/>
                </Border>
                <ContentPresenter VerticalAlignment="Center"/>
              </StackPanel>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="Bd"        Property="Background"   Value="#5BB8D4"/>
                <Setter TargetName="Bd"        Property="BorderBrush"  Value="#5BB8D4"/>
                <Setter TargetName="Check"     Property="Background"   Value="#0A1C26"/>
                <Setter TargetName="Check"     Property="BorderBrush"  Value="#0A1C26"/>
                <Setter TargetName="CheckMark" Property="Visibility"   Value="Visible"/>
                <Setter Property="Foreground"  Value="#0A1C26"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="DataGrid">
      <Setter Property="Background"               Value="#081520"/>
      <Setter Property="Foreground"               Value="#C8E4EF"/>
      <Setter Property="BorderBrush"              Value="#1A3D52"/>
      <Setter Property="BorderThickness"          Value="1"/>
      <Setter Property="GridLinesVisibility"      Value="Horizontal"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#132C3D"/>
      <Setter Property="RowBackground"            Value="#081520"/>
      <Setter Property="AlternatingRowBackground" Value="#0A1C2A"/>
      <Setter Property="ColumnHeaderHeight"       Value="30"/>
      <Setter Property="RowHeight"                Value="26"/>
      <Setter Property="FontSize"                 Value="11"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Setter Property="Background" Value="#081520"/>
      <Setter Property="Foreground" Value="#C8E4EF"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#0F2535"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#1A3D52"/>
          <Setter Property="Foreground" Value="#F0F8FC"/>
        </Trigger>
        <MultiTrigger>
          <MultiTrigger.Conditions>
            <Condition Property="IsSelected"   Value="True"/>
            <Condition Property="IsMouseOver"  Value="True"/>
          </MultiTrigger.Conditions>
          <Setter Property="Background" Value="#1E4D66"/>
        </MultiTrigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background"      Value="#0F2535"/>
      <Setter Property="Foreground"      Value="#5BB8D4"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="FontSize"        Value="11"/>
      <Setter Property="Padding"         Value="10,0"/>
      <Setter Property="BorderBrush"     Value="#1A3D52"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"         Value="8,0"/>
      <Setter Property="Foreground"      Value="#C8E4EF"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="Transparent"/>
          <Setter Property="Foreground" Value="#F0F8FC"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" Background="#0F2535" Padding="20,10">
      <DockPanel>
        <StackPanel DockPanel.Dock="Left" Orientation="Horizontal" VerticalAlignment="Center">
          <Image x:Name="ImgLogo" Width="80" Height="80" Margin="0,0,16,0"
                 Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality"/>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="CREAR INCIDENCIA" FontSize="18" FontWeight="Bold"
                       Foreground="#C8E4EF"/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
              <Border Background="#1A4F6A" CornerRadius="8" Padding="8,2">
                <TextBlock x:Name="TxtHeaderTitle" Foreground="#5BB8D4"
                           FontSize="11" FontWeight="Bold" Text=""/>
              </Border>
            </StackPanel>
          </StackPanel>
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- Body -->
    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="420"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- Panel izquierdo -->
      <Border Grid.Column="0" Background="#091820" BorderBrush="#1A3D52"
              BorderThickness="0,0,1,0" Padding="16">
        <DockPanel>
          <TextBlock DockPanel.Dock="Top" Text="CAPTURA DE VISTA"
                     Style="{StaticResource Label}" Margin="0,0,0,8"/>
          <Border x:Name="ScreenshotPlaceholder" DockPanel.Dock="Top"
                  Background="#0F2535" CornerRadius="8" Height="220" Margin="0,0,0,10">
            <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
              <TextBlock Text="&#128444;" FontSize="40" HorizontalAlignment="Center" Foreground="#1A3D52"/>
              <TextBlock Text="Sin captura" FontSize="12" Foreground="#2A5570"
                         HorizontalAlignment="Center" Margin="0,6,0,0"/>
            </StackPanel>
          </Border>
          <Image x:Name="ImgScreenshot" DockPanel.Dock="Top"
                 Height="220" Margin="0,0,0,10"
                 Stretch="Uniform" StretchDirection="DownOnly"
                 Visibility="Collapsed"/>
          <TextBlock DockPanel.Dock="Top" Text="VIEWPOINT" Style="{StaticResource Label}"/>
          <Border DockPanel.Dock="Top" Background="#0F2535" CornerRadius="6"
                  Padding="12,8" Margin="0,0,0,14">
            <StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                <TextBlock Text="Vista: " Foreground="#4A8BA6" FontSize="11"/>
                <TextBlock x:Name="TxtVista" Text="&#8212;" Foreground="#C8E4EF" FontSize="11"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                <TextBlock Text="Tipo: " Foreground="#4A8BA6" FontSize="11"/>
                <TextBlock x:Name="TxtTipoVista" Text="&#8212;" Foreground="#C8E4EF" FontSize="11"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="Camara: " Foreground="#4A8BA6" FontSize="11"/>
                <TextBlock x:Name="TxtCamara" Text="&#8212;" Foreground="#8ABCCE" FontSize="10"
                           TextWrapping="Wrap"/>
              </StackPanel>
            </StackPanel>
          </Border>
          <TextBlock DockPanel.Dock="Top" Text="ELEMENTOS SELECCIONADOS"
                     Style="{StaticResource Label}"/>
          <Button x:Name="BtnSeleccionar" DockPanel.Dock="Top"
                  Content="Seleccionar en Modelo"
                  Style="{StaticResource BtnGhost}"
                  HorizontalAlignment="Stretch" Margin="0,0,0,8"/>
          <DataGrid x:Name="GridElementos" AutoGenerateColumns="False"
                    IsReadOnly="True" SelectionMode="Extended"
                    HorizontalScrollBarVisibility="Auto"
                    VerticalScrollBarVisibility="Auto">
            <DataGrid.Columns>
              <DataGridTextColumn Header="ID"        Binding="{Binding ElementId}" Width="70"/>
              <DataGridTextColumn Header="Categoria" Binding="{Binding Categoria}" Width="120"/>
              <DataGridTextColumn Header="Tipo"      Binding="{Binding Tipo}"      Width="*"/>
            </DataGrid.Columns>
          </DataGrid>
        </DockPanel>
      </Border>

      <!-- Panel derecho -->
      <ScrollViewer Grid.Column="1" VerticalScrollBarVisibility="Auto" Padding="20,16">
        <StackPanel>
          <TextBlock Text="DETALLES DEL ISSUE" Style="{StaticResource Label}" Margin="0,0,0,4"/>
          <TextBlock Text="Titulo *" Style="{StaticResource Label}"/>
          <TextBox x:Name="TxtTitulo" Style="{StaticResource Input}"
                   Tag="Ej: Conflicto entre viga y ducto HVAC"/>
          <TextBlock Text="Descripcion" Style="{StaticResource Label}"/>
          <TextBox x:Name="TxtDescripcion" Style="{StaticResource Input}"
                   Height="100" TextWrapping="Wrap" AcceptsReturn="True"
                   VerticalScrollBarVisibility="Auto"
                   Tag="Describe el problema, ubicacion o accion requerida..."/>
          <TextBlock Text="Prioridad" Style="{StaticResource Label}"/>
          <ComboBox x:Name="CmbPrioridad" Style="{StaticResource Combo}">
            <ComboBox.ItemContainerStyle>
              <Style TargetType="ComboBoxItem">
                <Setter Property="Background"  Value="#0D2234"/>
                <Setter Property="Foreground"  Value="#F0F8FC"/>
                <Setter Property="Padding"     Value="10,8"/>
                <Setter Property="FontSize"    Value="13"/>
                <Style.Triggers>
                  <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#1A4F6A"/>
                  </Trigger>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1A4F6A"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </ComboBox.ItemContainerStyle>
            <ComboBoxItem Content="Critica"  Tag="Critica"/>
            <ComboBoxItem Content="Alta"     Tag="Alta" IsSelected="True"/>
            <ComboBoxItem Content="Media"    Tag="Media"/>
            <ComboBoxItem Content="Baja"     Tag="Baja"/>
          </ComboBox>

          <!-- ── Reportado por ── -->
          <TextBlock Text="Reportado por" Style="{StaticResource Label}"/>
          <ComboBox x:Name="CmbReportadoPor"
                    IsEditable="True"
                    Style="{StaticResource ComboPersona}">
            <ComboBox.ItemContainerStyle>
              <Style TargetType="ComboBoxItem">
                <Setter Property="Background"                Value="#0D2234"/>
                <Setter Property="Foreground"                Value="#F0F8FC"/>
                <Setter Property="Padding"                   Value="10,6"/>
                <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                <Style.Triggers>
                  <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#1A4F6A"/>
                  </Trigger>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1A4F6A"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </ComboBox.ItemContainerStyle>
            <ComboBox.ItemTemplate>
              <DataTemplate>
                <StackPanel Margin="0,3,0,4">
                  <TextBlock Text="{Binding Nombre}" FontSize="12"
                             Foreground="#F0F8FC" FontWeight="SemiBold"/>
                  <TextBlock Text="{Binding Email}" FontSize="10"
                             Foreground="#7DD4EC" Margin="0,3,0,0"/>
                </StackPanel>
              </DataTemplate>
            </ComboBox.ItemTemplate>
          </ComboBox>

          <!-- ── Asignado a ── -->
          <TextBlock Text="Asignado a" Style="{StaticResource Label}"/>
          <ComboBox x:Name="CmbAsignado"
                    IsEditable="True"
                    Style="{StaticResource ComboPersona}">
            <ComboBox.ItemContainerStyle>
              <Style TargetType="ComboBoxItem">
                <Setter Property="Background"                Value="#0D2234"/>
                <Setter Property="Foreground"                Value="#F0F8FC"/>
                <Setter Property="Padding"                   Value="10,6"/>
                <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                <Style.Triggers>
                  <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#1A4F6A"/>
                  </Trigger>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1A4F6A"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </ComboBox.ItemContainerStyle>
            <ComboBox.ItemTemplate>
              <DataTemplate>
                <StackPanel Margin="0,3,0,4">
                  <TextBlock Text="{Binding Nombre}" FontSize="12"
                             Foreground="#F0F8FC" FontWeight="SemiBold"/>
                  <TextBlock Text="{Binding Email}" FontSize="10"
                             Foreground="#7DD4EC" Margin="0,3,0,0"/>
                </StackPanel>
              </DataTemplate>
            </ComboBox.ItemTemplate>
          </ComboBox>

          <Button x:Name="BtnGestionarPersonas"
                  Content="Gestionar directorio de personas..."
                  Style="{StaticResource BtnLink}" HorizontalAlignment="Left"
                  Margin="0,6,0,0"/>

          <TextBlock Text="Disciplinas involucradas" Style="{StaticResource Label}"/>
          <WrapPanel>
            <CheckBox x:Name="ChkArquitectura" Content="Arquitectura"
                      Style="{StaticResource ChkDisciplina}" IsChecked="False"/>
            <CheckBox x:Name="ChkEstructura"   Content="Estructura"
                      Style="{StaticResource ChkDisciplina}" IsChecked="False"/>
            <CheckBox x:Name="ChkMEP"          Content="MEP"
                      Style="{StaticResource ChkDisciplina}" IsChecked="False"/>
            <CheckBox x:Name="ChkCivil"        Content="Civil"
                      Style="{StaticResource ChkDisciplina}" IsChecked="False"/>
            <CheckBox x:Name="ChkGeneral"      Content="General"
                      Style="{StaticResource ChkDisciplina}" IsChecked="False"/>
          </WrapPanel>
          <!-- Archivos adjuntos -->
          <Border Height="1" Background="#1A3D52" Margin="0,14,0,14"/>
          <TextBlock Text="Archivos adjuntos" Style="{StaticResource Label}" Margin="0,0,0,6"/>
          <Border Background="#0F2535" CornerRadius="6" Padding="10,8">
            <StackPanel>
              <TextBlock x:Name="TxtSinAdjuntos" Text="Sin archivos adjuntos"
                         FontSize="11" Foreground="#2A5570"/>
              <ListBox x:Name="LstAdjuntos" Background="Transparent" BorderThickness="0"
                       MaxHeight="90" DisplayMemberPath="Nombre" Visibility="Collapsed"
                       ScrollViewer.VerticalScrollBarVisibility="Auto">
                <ListBox.ItemContainerStyle>
                  <Style TargetType="ListBoxItem">
                    <Setter Property="Foreground" Value="#C8E4EF"/>
                    <Setter Property="FontSize"   Value="11"/>
                    <Setter Property="Background" Value="Transparent"/>
                    <Setter Property="Padding"    Value="4,3"/>
                    <Style.Triggers>
                      <Trigger Property="IsSelected" Value="True">
                        <Setter Property="Background" Value="#1A3D52"/>
                      </Trigger>
                    </Style.Triggers>
                  </Style>
                </ListBox.ItemContainerStyle>
              </ListBox>
              <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                <Button x:Name="BtnAdjuntar" Content="+ Adjuntar archivo"
                        Style="{StaticResource BtnGhost}" Margin="0,0,8,0"/>
                <Button x:Name="BtnQuitarAdjunto" Content="Quitar seleccionado"
                        Style="{StaticResource BtnGhost}" IsEnabled="False"/>
              </StackPanel>
            </StackPanel>
          </Border>

          <Border Height="1" Background="#1A3D52" Margin="0,16,0,16"/>
          <TextBlock Text="PROYECTO" Style="{StaticResource Label}" Margin="0,0,0,6"/>
          <Border Background="#0F2535" CornerRadius="6" Padding="12,10">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <TextBlock Grid.Row="0" Grid.Column="0" Text="Nombre:" Foreground="#4A8BA6"
                         FontSize="11" Margin="0,0,12,6"/>
              <TextBlock Grid.Row="0" Grid.Column="1" x:Name="TxtProyecto"
                         Text="&#8212;" Foreground="#C8E4EF" FontSize="11" Margin="0,0,0,6"/>
              <TextBlock Grid.Row="1" Grid.Column="0" Text="Fecha:" Foreground="#4A8BA6"
                         FontSize="11" Margin="0,0,12,6"/>
              <TextBlock Grid.Row="1" Grid.Column="1" x:Name="TxtFecha"
                         Text="&#8212;" Foreground="#C8E4EF" FontSize="11" Margin="0,0,0,6"/>
              <TextBlock Grid.Row="2" Grid.Column="0" Text="Autor:" Foreground="#4A8BA6"
                         FontSize="11" Margin="0,0,12,0"/>
              <TextBlock Grid.Row="2" Grid.Column="1" x:Name="TxtAutor"
                         Text="&#8212;" Foreground="#C8E4EF" FontSize="11"/>
            </Grid>
          </Border>
          <!-- ── Acciones ── -->
          <Border Height="1" Background="#1A3D52" Margin="0,20,0,16"/>
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Right"
                      Margin="0,0,0,8">
            <Button x:Name="BtnManual" Content="Manual"
                    Style="{StaticResource BtnGhost}" Margin="0,0,10,0"
                    Background="#2A5C3D" ToolTip="Abrir manual de usuario"/>
            <Button x:Name="BtnCancelar" Content="Cancelar"
                    Style="{StaticResource BtnGhost}"
                    Margin="0,0,10,0"/>
            <Button x:Name="BtnCrear" Content="Crear Incidencia"
                    Style="{StaticResource BtnGhost}"/>
          </StackPanel>
        </StackPanel>
      </ScrollViewer>
    </Grid>

    <!-- Footer: solo estado -->
    <Border Grid.Row="2" Background="#0A1C26" BorderBrush="#1A3D52"
            BorderThickness="0,1,0,0" Padding="20,10">
      <TextBlock x:Name="TxtEstado"
                 FontSize="11" Foreground="#4A8BA6" VerticalAlignment="Center"
                 Text="Completa los campos y captura la vista antes de guardar."/>
    </Border>
  </Grid>
</Window>
"""

# ── Modelos de datos ────────────────────────────────────────────────────────────
class ElementoInfo(object):
    def __init__(self, eid, categoria, tipo):
        self.ElementId = eid
        self.Categoria = categoria
        self.Tipo      = tipo


class AdjuntoItem(object):
    """Representa un archivo adjunto a la incidencia."""
    def __init__(self, nombre, ruta):
        self.Nombre = nombre   # nombre del archivo (basename)
        self.Ruta   = ruta     # ruta completa al archivo de origen

    def ToString(self):
        return self.Nombre


# ── Helpers ────────────────────────────────────────────────────────────────────
def _get_type_name(element):
    try:
        t = doc.GetElement(element.GetTypeId())
        if t:
            return t.Name or "—"
    except Exception:
        pass
    return "—"


def _get_viewpoint_info(view):
    info = {
        "view_name": view.Name,
        "view_type": str(view.ViewType),
        "view_id":   view.Id.IntegerValue,   # ID persistente de la vista en el documento
        "camera":    None,
    }
    try:
        if isinstance(view, View3D):
            ori = view.GetOrientation()
            eye = ori.EyePosition
            fwd = ori.ForwardDirection
            up  = ori.UpDirection
            info["camera"] = {
                "eye":     [round(eye.X, 4), round(eye.Y, 4), round(eye.Z, 4)],
                "forward": [round(fwd.X, 4), round(fwd.Y, 4), round(fwd.Z, 4)],
                "up":      [round(up.X,  4), round(up.Y,  4), round(up.Z,  4)],
            }
            # Guardar crop box para restaurar el zoom exacto de la vista
            # Min/Max estan en coordenadas locales de la vista (sistema definido por la orientacion)
            try:
                cb = view.CropBox
                info["camera"]["crop_box"] = {
                    "active": bool(view.CropBoxActive),
                    "min": [round(cb.Min.X, 6), round(cb.Min.Y, 6), round(cb.Min.Z, 6)],
                    "max": [round(cb.Max.X, 6), round(cb.Max.Y, 6), round(cb.Max.Z, 6)],
                }
            except Exception:
                pass
    except Exception:
        pass
    return info


def _capture_view(view):
    """
    Captura únicamente el área de la vista activa en pantalla
    (sin ribbon, paneles ni paletas) usando UIView.GetWindowRectangle()
    + System.Drawing.Graphics.CopyFromScreen.
    """
    tmp_dir  = tempfile.mkdtemp(prefix="bimissue_")
    png_path = os.path.join(tmp_dir, "screenshot.png")

    x = y = width = height = None
    try:
        for uiv in uidoc.GetOpenUIViews():
            if uiv.ViewId == view.Id:
                r      = uiv.GetWindowRectangle()
                x      = r.Left
                y      = r.Top
                width  = r.Right  - r.Left
                height = r.Bottom - r.Top
                break
    except Exception:
        pass

    if not width or not height:
        import ctypes

        hwnd = __revit__.MainWindowHandle.ToInt32()   # noqa: F821

        class RECT(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top",  ctypes.c_long),
                        ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

        rect = RECT()
        if hwnd and ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            x, y  = rect.left, rect.top
            width  = rect.right  - rect.left
            height = rect.bottom - rect.top
        else:
            b = Screen.PrimaryScreen.Bounds
            x, y, width, height = b.X, b.Y, b.Width, b.Height

    bmp = Bitmap(width, height)
    gfx = Graphics.FromImage(bmp)
    gfx.CopyFromScreen(x, y, 0, 0, DSize(width, height))
    bmp.Save(png_path, ImageFormat.Png)
    gfx.Dispose()
    bmp.Dispose()

    return png_path


def _escape_html(s):
    """Escapa caracteres especiales para HTML."""
    return (s or u"").replace(u"&", u"&amp;").replace(u"<", u"&lt;").replace(u">", u"&gt;").replace(u'"', u"&quot;")


def _get_image_dimensions_bimissue(path):
    """Obtiene ancho y alto de PNG. Retorna (width, height) o (None, None)."""
    try:
        import struct
        with open(path, "rb") as f:
            header = f.read(24)
        if header[:8] == b"\x89PNG\r\n\x1a\n" and header[12:16] == b"IHDR":
            w, h = struct.unpack(">II", header[16:24])
            return w, h
    except Exception:
        pass
    return None, None


def _load_logo_base64_bimissue():
    """Carga logo.png desde Recordatorios. Retorna (data_uri, width, height) escalado a max 50px alto."""
    try:
        import base64
        try:
            _script_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            try:
                _script_dir = os.path.dirname(os.path.abspath(__commandpath__))
            except NameError:
                return u"", 180, 50
        _recordatorios = os.path.join(os.path.dirname(_script_dir), u"Recordatorios")
        _logo_path = os.path.join(_recordatorios, u"logo.png")
        if not os.path.isfile(_logo_path):
            return u"", 180, 50
        with open(_logo_path, "rb") as f:
            data = base64.b64encode(f.read())
        try:
            data_str = data.decode("ascii")
        except Exception:
            return u"", 180, 50
        data_uri = u"data:image/png;base64,{}".format(data_str)
        orig_w, orig_h = _get_image_dimensions_bimissue(_logo_path)
        if orig_w and orig_h and orig_h > 0:
            display_h = 50
            display_w = int(orig_w * display_h / orig_h)
            if display_w > 300:
                display_w = 300
                display_h = int(orig_h * display_w / orig_w)
            return data_uri, display_w, display_h
        return data_uri, 180, 50
    except Exception:
        return u"", 180, 50


def _build_email_html_bimissue(asi_nom, proyecto, issue_number, titulo, descripcion, logo_data_uri=u"", logo_w=180, logo_h=50):
    """Construye el cuerpo HTML del correo de incidencia."""
    _link_revisar = u"file:///Y:/00_SERVIDOR%20DE%20INCIDENCIAS/ConsultaIncidencias/Abrir.bat"
    parts = []
    parts.append(u"<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#333;max-width:600px'>")
    if logo_data_uri:
        parts.append(u"<div style='background:" + _COLOR_WHITE + u";padding:20px 24px;text-align:left;border-bottom:3px solid " + _COLOR_LIGHT + u"'>")
        # Dimensiones proporcionales para Outlook clasico y nuevo
        parts.append(u"<img src='" + logo_data_uri + u"' alt='Arainco' width='" + str(logo_w) + u"' height='" + str(logo_h) + u"' style='display:block;width:" + str(logo_w) + u"px;height:" + str(logo_h) + u"px;border:0' />")
        parts.append(u"</div>")
    else:
        parts.append(u"<div style='background:" + _COLOR_DARK + u";padding:16px 24px;border-bottom:3px solid " + _COLOR_LIGHT + u"'>")
        parts.append(u"<span style='color:" + _COLOR_WHITE + u";font-size:20px;font-weight:bold'>ARAINCO</span> ")
        parts.append(u"<span style='color:" + _COLOR_LIGHT + u";font-size:12px'>INGENIERIA ESTRUCTURAL</span>")
        parts.append(u"</div>")
    parts.append(u"<div style='padding:24px;background:" + _COLOR_WHITE + u"'>")
    parts.append(u"<p style='color:" + _COLOR_DARK + u";margin:0 0 16px'>Estimado/a " + _escape_html(asi_nom or u"colega") + u",</p>")
    proy_esc = _escape_html(proyecto or u"-")
    parts.append(u"<p style='color:#555;margin:0 0 20px'>Se te ha asignado una nueva incidencia en el proyecto " + proy_esc + u".</p>")
    parts.append(u"<div style='margin-bottom:16px;padding:12px 16px;background:#f8fafc;border-left:4px solid " + _COLOR_LIGHT + u";border-radius:4px'>")
    parts.append(u"<div style='color:" + _COLOR_DARK + u";font-weight:bold;font-size:13px;margin-bottom:8px'>Incidencia N" + str(issue_number) + u"</div>")
    parts.append(u"<div style='color:#444;font-size:13px;padding:4px 0;border-bottom:1px solid " + _COLOR_GRAY + u"'>" + _escape_html(titulo or u"") + u"</div>")
    desc_br = _escape_html((descripcion or u"").replace(u"\n", u"<br>"))
    parts.append(u"<div style='color:#444;font-size:13px;padding:4px 0'>" + desc_br + u"</div>")
    parts.append(u"</div>")
    parts.append(u"<p style='color:#666;margin:24px 0 0;font-size:13px'>Por favor, revisar la aplicacion <a href='" + _link_revisar + u"' style='color:" + _COLOR_LIGHT + u";text-decoration:underline'>revisar incidencias</a> para mas detalles.</p>")
    parts.append(u"</div>")
    parts.append(u"<div style='padding:12px 24px;background:" + _COLOR_DARK + u";color:" + _COLOR_WHITE + u";font-size:11px;border-top:1px solid " + _COLOR_LIGHT + u"'>")
    parts.append(u"Arainco Ingenieria Estructural - Notificacion de incidencia</div>")
    parts.append(u"</div>")
    return u"".join(parts)


def _send_mailto_bimissue(to_email, subject, body_plain, body_html, from_email=None):
    """
    Abre el cliente de correo con To, Subject y Body (HTML + plain).
    Usa archivo .eml porque el nuevo Outlook ignora los parametros de mailto (bug conocido).
    """
    try:
        import tempfile
        try:
            from email.mime.text import MIMEText
            from email.mime.multipart import MIMEMultipart
            from email.generator import Generator
        except ImportError:
            return False

        msg = MIMEMultipart("alternative")
        msg["To"] = to_email
        msg["Subject"] = subject or u""
        msg["From"] = (from_email or u"noreply@local").strip()
        msg["X-Unsent"] = "1"
        msg.attach(MIMEText((body_plain or u"").replace(u"\r\n", u"\n"), "plain", "utf-8"))
        msg.attach(MIMEText(body_html or body_plain or u"", "html", "utf-8"))

        eml_path = os.path.join(tempfile.gettempdir(), u"incidencia_{}.eml".format(os.getpid()))
        with open(eml_path, "wb") as f:
            gen = Generator(f, mangle_from_=False)
            gen.flatten(msg)

        os.startfile(eml_path)
        return True
    except Exception:
        return False


def _save_issue(issue_dir, screenshot_src, issue_data, elementos):
    import os
    import json
    # Deriva la carpeta raíz (padre) de issue_dir para crearla si no existe
    issues_root = os.path.dirname(issue_dir)
    if not sio.Directory.Exists(issues_root):
        sio.Directory.CreateDirectory(issues_root)
    sio.Directory.CreateDirectory(issue_dir)

    png_dst = os.path.join(issue_dir, "screenshot.png")
    if screenshot_src and sio.File.Exists(screenshot_src):
        sio.File.Copy(screenshot_src, png_dst, True)

    issue_data["elementos"] = [
        {"id": e.ElementId, "categoria": e.Categoria, "tipo": e.Tipo}
        for e in elementos
    ]
    json_path = os.path.join(issue_dir, "issue.json")
    sio.File.WriteAllText(
        json_path,
        json.dumps(issue_data, ensure_ascii=False, indent=2),
        System.Text.Encoding.UTF8
    )
    return png_dst, json_path


# ── Diálogo Gestionar Personas ──────────────────────────────────────────────
class GestionarPersonasDialog(object):
    """
    Ventana modal para agregar y eliminar personas del directorio.
    Trabaja directamente sobre el ObservableCollection compartido.
    """

    def __init__(self, personas_collection, issues_dir, personas_file, owner=None):
        self._personas      = personas_collection
        self._issues_dir    = issues_dir
        self._personas_file = personas_file
        win = XamlReader.Parse(GESTIONAR_PERSONAS_XAML)
        self._win = win

        self._txt_nombre = win.FindName("TxtNombreNew")
        self._txt_email  = win.FindName("TxtEmailNew")
        self._btn_add    = win.FindName("BtnAgregar")
        self._grid       = win.FindName("GridPersonas")
        self._btn_del    = win.FindName("BtnEliminar")
        self._btn_close  = win.FindName("BtnCerrar")

        self._grid.ItemsSource = self._personas

        self._btn_add.Click   += RoutedEventHandler(self._on_agregar)
        self._btn_del.Click   += RoutedEventHandler(self._on_eliminar)
        self._btn_close.Click += RoutedEventHandler(lambda s, e: win.Close())

        if owner:
            try:
                win.Owner = owner
            except Exception:
                pass

        win.ShowDialog()

    def _on_agregar(self, sender, args):
        nombre = self._txt_nombre.Text.strip()
        if not nombre:
            return
        email = self._txt_email.Text.strip()
        self._personas.Add(PersonaItem(nombre, email))
        _save_personas(list(self._personas), self._issues_dir, self._personas_file)
        self._txt_nombre.Text = ""
        self._txt_email.Text  = ""
        self._txt_nombre.Focus()

    def _on_eliminar(self, sender, args):
        sel = self._grid.SelectedItem
        if sel:
            self._personas.Remove(sel)
            _save_personas(list(self._personas), self._issues_dir, self._personas_file)


# ── Ventana principal ───────────────────────────────────────────────────────────
class BIMIssueWindow(object):
    """
    Wrapper sobre la ventana WPF cargada con XamlReader.Parse.
    Los eventos se registran directamente sobre los elementos del árbol visual.
    """

    def __init__(self, initial_png=None, initial_view=None):
        # En pyRevit, los handlers WPF no acceden al scope global del módulo.
        # Todo lo necesario se almacena como atributo de instancia aquí,
        # donde el scope de módulo SÍ está disponible.
        self._issues_dir    = ISSUES_DIR
        self._personas_file = PERSONAS_FILE

        # Carpeta del proyecto: NombreArchivoRevit (central si es colaborativo)
        _project_folder = _get_project_folder_name(doc)
        self._project_issues_dir = os.path.join(ISSUES_DIR, _project_folder)

        # Referencias a clases/funciones del módulo usadas en handlers
        self._PersonaItem   = PersonaItem
        self._ElementoInfo  = ElementoInfo
        self._AdjuntoItem   = AdjuntoItem
        self._GestionarPersonasDialog = GestionarPersonasDialog
        self._get_type_name_fn = _get_type_name
        self._get_viewpoint_info_fn = _get_viewpoint_info
        self._OC = ObservableCollection  # para crear colecciones filtradas en handlers
        self._updating_combos = False    # evita bucles al actualizar ItemsSource

        win = XamlReader.Parse(XAML)
        self._win = win

        self._img            = win.FindName("ImgScreenshot")
        self._ph             = win.FindName("ScreenshotPlaceholder")
        self._txt_vista      = win.FindName("TxtVista")
        self._txt_tipo       = win.FindName("TxtTipoVista")
        self._txt_cam        = win.FindName("TxtCamara")
        self._grid           = win.FindName("GridElementos")
        self._titulo         = win.FindName("TxtTitulo")
        self._desc           = win.FindName("TxtDescripcion")
        self._prioridad      = win.FindName("CmbPrioridad")
        self._cmb_reportado  = win.FindName("CmbReportadoPor")
        self._cmb_asignado   = win.FindName("CmbAsignado")
        self._btn_gestionar  = win.FindName("BtnGestionarPersonas")
        self._chk_arquitectura = win.FindName("ChkArquitectura")
        self._chk_estructura   = win.FindName("ChkEstructura")
        self._chk_mep          = win.FindName("ChkMEP")
        self._chk_civil        = win.FindName("ChkCivil")
        self._chk_general      = win.FindName("ChkGeneral")
        self._proyecto       = win.FindName("TxtProyecto")
        self._fecha          = win.FindName("TxtFecha")
        self._autor          = win.FindName("TxtAutor")
        self._estado_lbl     = win.FindName("TxtEstado")
        self._lst_adjuntos       = win.FindName("LstAdjuntos")
        self._txt_sin_adjuntos   = win.FindName("TxtSinAdjuntos")
        self._btn_quitar_adjunto = win.FindName("BtnQuitarAdjunto")

        self._screenshot_path = None
        self._viewpoint_info  = {}
        self._elementos = ObservableCollection[object]()
        self._grid.ItemsSource = self._elementos
        self._adjuntos_col = ObservableCollection[object]()
        self._lst_adjuntos.ItemsSource = self._adjuntos_col

        self._personas = ObservableCollection[object]()
        for p in _load_personas(self._personas_file):
            self._personas.Add(p)
        self._cmb_reportado.ItemsSource = sorted(list(self._personas), key=lambda p: (p.Nombre or "").lower())
        self._cmb_asignado.ItemsSource  = sorted(list(self._personas), key=lambda p: (p.Nombre or "").lower())
        # Sin texto libre: IsReadOnly=True en el template, solo selección de lista

        self._fill_project_info()
        self._load_logo()
        self._set_issue_number()

        win.FindName("BtnSeleccionar").Click += RoutedEventHandler(self._on_seleccionar)
        win.FindName("BtnCancelar").Click    += RoutedEventHandler(self._on_cancelar)
        win.FindName("BtnCrear").Click       += RoutedEventHandler(self._on_guardar)
        win.FindName("BtnManual").Click      += RoutedEventHandler(self._on_manual)
        self._btn_gestionar.Click            += RoutedEventHandler(self._on_gestionar_personas)
        win.FindName("BtnAdjuntar").Click    += RoutedEventHandler(self._on_adjuntar)
        self._btn_quitar_adjunto.Click       += RoutedEventHandler(self._on_quitar_adjunto)
        self._lst_adjuntos.SelectionChanged  += SelectionChangedEventHandler(self._on_adjunto_selected)

        self._cmb_reportado.SelectionChanged += SelectionChangedEventHandler(self._on_reportado_changed)
        self._cmb_asignado.SelectionChanged  += SelectionChangedEventHandler(self._on_asignado_changed)
        self._win.KeyDown                    += KeyEventHandler(self._on_key_down)

        if initial_png and initial_view:
            self._load_screenshot(initial_png, initial_view)

    def _set_estado(self, msg):
        try:
            if self._estado_lbl:
                self._estado_lbl.Text = msg
        except Exception:
            pass

    def _set_issue_number(self):
        try:
            import os
            existing = 0
            # Cuenta issues dentro de la carpeta del proyecto
            if os.path.isdir(self._project_issues_dir):
                existing = sum(
                    1 for n in os.listdir(self._project_issues_dir)
                    if n.startswith("ISSUE_") and
                    os.path.isdir(os.path.join(self._project_issues_dir, n))
                )
            lbl = self._win.FindName("TxtHeaderTitle")
            if lbl:
                lbl.Text = u"Incidencia N\u00ba {}".format(existing + 1)
        except Exception:
            pass

    def _load_logo(self):
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
            if not os.path.exists(logo_path):
                return
            img_ctrl = self._win.FindName("ImgLogo")
            if not img_ctrl:
                return
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource   = Uri(logo_path, UriKind.Absolute)
            bmp.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad
            bmp.EndInit()
            bmp.Freeze()
            img_ctrl.Source = bmp
        except Exception:
            pass

    def _fill_project_info(self):
        import datetime
        # Nombre del proyecto: directorio que contiene el modelo central (o el archivo actual)
        self._proyecto.Text = _get_project_folder_name(doc) or u"—"
        try:
            pi = doc.ProjectInformation
            self._autor.Text = pi.Author or u"—"
        except Exception:
            try:
                self._autor.Text = u"—"
            except Exception:
                pass
        try:
            self._fecha.Text = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass

    def _load_screenshot(self, png, view):
        try:
            self._screenshot_path = png
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource   = Uri(png, UriKind.Absolute)
            bmp.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad
            bmp.EndInit()
            bmp.Freeze()
            self._img.Source     = bmp
            self._img.Visibility = System.Windows.Visibility.Visible
            self._ph.Visibility  = System.Windows.Visibility.Collapsed
            vp = self._get_viewpoint_info_fn(view)
            self._viewpoint_info = vp
            self._txt_vista.Text = vp["view_name"]
            self._txt_tipo.Text  = vp["view_type"]
            cam = vp.get("camera")
            self._txt_cam.Text = (
                u"Eye: ({}, {}, {})".format(*cam["eye"]) if cam
                else u"Vista 2D / sin datos de camara"
            )
            self._set_estado(u"Vista capturada: {}".format(view.Name))
        except Exception:
            pass

    def _on_reportado_changed(self, sender, args):
        if self._updating_combos:
            return
        self._updating_combos = True
        try:
            sel = self._cmb_reportado.SelectedItem
            if sel and isinstance(sel, self._PersonaItem):
                filtered = [p for p in self._personas if p.Nombre != sel.Nombre]
                # Si Asignado tiene al mismo seleccionado, limpiar
                if (self._cmb_asignado.SelectedItem and
                        isinstance(self._cmb_asignado.SelectedItem, self._PersonaItem) and
                        self._cmb_asignado.SelectedItem.Nombre == sel.Nombre):
                    self._cmb_asignado.SelectedItem = None
                    self._cmb_asignado.Text = u""
                self._cmb_asignado.ItemsSource = sorted(filtered, key=lambda p: (p.Nombre or "").lower())
            else:
                # Sin selección: restaurar lista completa en Asignado
                _excluir = None
                if (self._cmb_asignado.SelectedItem and
                        isinstance(self._cmb_asignado.SelectedItem, self._PersonaItem)):
                    _excluir = self._cmb_asignado.SelectedItem.Nombre
                filtered = [p for p in self._personas if p.Nombre != _excluir]
                self._cmb_reportado.ItemsSource = sorted(filtered, key=lambda p: (p.Nombre or "").lower())
        except Exception:
            pass
        finally:
            self._updating_combos = False

    def _on_asignado_changed(self, sender, args):
        if self._updating_combos:
            return
        self._updating_combos = True
        try:
            sel = self._cmb_asignado.SelectedItem
            if sel and isinstance(sel, self._PersonaItem):
                filtered = [p for p in self._personas if p.Nombre != sel.Nombre]
                # Si Reportado tiene al mismo seleccionado, limpiar
                if (self._cmb_reportado.SelectedItem and
                        isinstance(self._cmb_reportado.SelectedItem, self._PersonaItem) and
                        self._cmb_reportado.SelectedItem.Nombre == sel.Nombre):
                    self._cmb_reportado.SelectedItem = None
                    self._cmb_reportado.Text = u""
                self._cmb_reportado.ItemsSource = sorted(filtered, key=lambda p: (p.Nombre or "").lower())
            else:
                # Sin selección: restaurar lista completa en Reportado
                _excluir = None
                if (self._cmb_reportado.SelectedItem and
                        isinstance(self._cmb_reportado.SelectedItem, self._PersonaItem)):
                    _excluir = self._cmb_reportado.SelectedItem.Nombre
                filtered = [p for p in self._personas if p.Nombre != _excluir]
                self._cmb_asignado.ItemsSource = sorted(filtered, key=lambda p: (p.Nombre or "").lower())
        except Exception:
            pass
        finally:
            self._updating_combos = False

    # -- Adjuntos --------------------------------------------------------------
    def _on_adjuntar(self, sender, args):
        try:
            import clr as _clr
            _clr.AddReference("System.Windows.Forms")
            from System.Windows.Forms import OpenFileDialog, DialogResult as DR
            import os as _os
            import System.Windows as _sw

            dlg = OpenFileDialog()
            dlg.Title      = u"Seleccionar archivos adjuntos"
            dlg.Filter     = (
                u"Todos los archivos (*.*)|*.*"
                u"|PDF (*.pdf)|*.pdf"
                u"|Word (*.docx;*.doc)|*.docx;*.doc"
                u"|Excel (*.xlsx;*.xls)|*.xlsx;*.xls"
                u"|Imagenes (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg"
            )
            dlg.FilterIndex  = 1
            dlg.Multiselect  = True

            if dlg.ShowDialog() != DR.OK:
                return

            _Adj = self._AdjuntoItem
            existing = set(a.Nombre for a in list(self._adjuntos_col))
            for ruta in dlg.FileNames:
                nombre = _os.path.basename(ruta)
                if nombre not in existing:
                    self._adjuntos_col.Add(_Adj(nombre, ruta))
                    existing.add(nombre)
            self._refresh_adjuntos_ui()
        except Exception as ex:
            self._set_estado(u"Error al adjuntar archivo: " + str(ex))

    def _on_quitar_adjunto(self, sender, args):
        try:
            sel = self._lst_adjuntos.SelectedItem
            if sel:
                self._adjuntos_col.Remove(sel)
                self._refresh_adjuntos_ui()
        except Exception:
            pass

    def _on_adjunto_selected(self, sender, args):
        try:
            self._btn_quitar_adjunto.IsEnabled = (
                self._lst_adjuntos.SelectedItem is not None
            )
        except Exception:
            pass

    def _refresh_adjuntos_ui(self):
        try:
            import System.Windows as _sw
            count = self._adjuntos_col.Count
            self._txt_sin_adjuntos.Visibility = (
                _sw.Visibility.Collapsed if count > 0 else _sw.Visibility.Visible
            )
            self._lst_adjuntos.Visibility = (
                _sw.Visibility.Visible if count > 0 else _sw.Visibility.Collapsed
            )
            if count == 0:
                self._btn_quitar_adjunto.IsEnabled = False
        except Exception:
            pass

    # -- Seleccion en modelo ---------------------------------------------------
    def _on_seleccionar(self, sender, args):
        self._win.Hide()
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                u"Selecciona elementos y pulsa ENTER para confirmar"
            )
            if refs:
                ids_ok = {item.ElementId for item in self._elementos}
                nuevos = 0
                for ref in refs:
                    elem = doc.GetElement(ref.ElementId)
                    if not elem:
                        continue
                    eid = str(elem.Id.IntegerValue)
                    if eid in ids_ok:
                        continue
                    try:
                        cat = elem.Category.Name if elem.Category else u"—"
                    except Exception:
                        cat = u"—"
                    self._elementos.Add(
                        self._ElementoInfo(eid, cat, self._get_type_name_fn(elem))
                    )
                    ids_ok.add(eid)
                    nuevos += 1
                self._set_estado(
                    u"Se agregaron {} elemento(s). Total: {}.".format(nuevos, len(self._elementos))
                )
        except Exception:
            self._set_estado(u"Seleccion cancelada.")
        finally:
            self._win.Show()

    def _on_cancelar(self, sender, args):
        self._win.Close()

    def _on_key_down(self, sender, args):
        try:
            if args.Key == Key.Escape:
                self._win.Close()
        except Exception:
            pass

    def _on_gestionar_personas(self, sender, args):
        self._GestionarPersonasDialog(
            self._personas,
            issues_dir=self._issues_dir,
            personas_file=self._personas_file,
            owner=self._win
        )
        # Refrescar combos con lista ordenada tras cerrar el diálogo
        _sorted = sorted(list(self._personas), key=lambda p: (p.Nombre or "").lower())
        self._cmb_reportado.ItemsSource = _sorted
        self._cmb_asignado.ItemsSource  = _sorted

    def _on_manual(self, sender, args):
        """Abre el manual de usuario en el navegador predeterminado."""
        try:
            _script_dir = os.path.dirname(os.path.abspath(__file__))
            manual_path = os.path.join(_script_dir, "manual_usuario.html")
            if os.path.exists(manual_path):
                abs_path = os.path.abspath(manual_path)
                os.startfile(abs_path)
                self._set_estado(u"Manual abierto.")
            else:
                MessageBox.Show(
                    u"No se encontró el archivo manual_usuario.html.\n\nRuta esperada:\n{}".format(manual_path),
                    u"Manual no encontrado",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                )
        except Exception as ex:
            self._set_estado(u"Error al abrir manual: " + str(ex))

    def _on_guardar(self, sender, args):
        import os
        import datetime
        import json
        import traceback as _tb
        _log = u""
        try:
            titulo = u""
            try:
                titulo = self._titulo.Text.strip()
            except Exception as _e:
                _log += u"[titulo err: {}] ".format(_e)

            if not titulo:
                self._set_estado(u"⚠ Campo obligatorio: Titulo.")
                try:
                    self._titulo.Focus()
                except Exception:
                    pass
                return

            descripcion = u""
            try:
                descripcion = self._desc.Text.strip()
            except Exception:
                pass
            if not descripcion:
                self._set_estado(u"⚠ Campo obligatorio: Descripcion.")
                try:
                    self._desc.Focus()
                except Exception:
                    pass
                return

            reportado = {"nombre": u"—", "email": u""}
            try:
                reportado = self._persona_from_combo(self._cmb_reportado)
            except Exception as _e:
                _log += u"[reportado err: {}] ".format(_e)
            if not reportado.get("nombre") or reportado["nombre"] == u"—":
                self._set_estado(u"⚠ Campo obligatorio: Reportado por.")
                try:
                    self._cmb_reportado.Focus()
                except Exception:
                    pass
                return

            asignado = {"nombre": u"—", "email": u""}
            try:
                asignado = self._persona_from_combo(self._cmb_asignado)
            except Exception as _e:
                _log += u"[asignado err: {}] ".format(_e)
            if not asignado.get("nombre") or asignado["nombre"] == u"—":
                self._set_estado(u"⚠ Campo obligatorio: Asignado a.")
                try:
                    self._cmb_asignado.Focus()
                except Exception:
                    pass
                return

            # Nota: envio hibrido (COM -> SMTP -> mailto) - ya no se exige Outlook abierto

            disciplinas = self._get_disciplinas()
            if not disciplinas or disciplinas == [u"General"]:
                # Comprobar que el usuario marcó al menos una disciplina explícitamente
                _alguna = False
                for _chk in [self._chk_arquitectura, self._chk_estructura,
                              self._chk_mep, self._chk_civil, self._chk_general]:
                    try:
                        if _chk and _chk.IsChecked:
                            _alguna = True
                            break
                    except Exception:
                        pass
                if not _alguna:
                    self._set_estado(u"⚠ Campo obligatorio: selecciona al menos una Disciplina.")
                    return

            prioridad = self._combo_tag(self._prioridad, u"Alta")
            estado    = u"Abierto"

            proyecto = u"—"
            try:
                proyecto = self._proyecto.Text
            except Exception:
                pass

            autor = u"—"
            try:
                autor = self._autor.Text
            except Exception:
                pass

            # Calcular número de incidencia (issues existentes en carpeta de proyecto + 1)
            _num_issues = 0
            if os.path.isdir(self._project_issues_dir):
                _num_issues = sum(
                    1 for _n in os.listdir(self._project_issues_dir)
                    if _n.startswith("ISSUE_") and
                    os.path.isdir(os.path.join(self._project_issues_dir, _n))
                )
            _issue_number = _num_issues + 1

            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            issue_data = {
                "id":               u"ISSUE_{}".format(ts),
                "numero":           _issue_number,
                "titulo":           titulo,
                "descripcion":      descripcion,
                "prioridad":        prioridad,
                "estado":           estado,
                "disciplinas":      disciplinas,
                "reportado_por":    reportado,
                "asignado_a":       asignado,
                "fecha":            datetime.datetime.now().isoformat(),
                "proyecto":         proyecto,
                "proyecto_carpeta": os.path.basename(self._project_issues_dir),
                "autor":            autor,
                "viewpoint":        self._viewpoint_info,
                "screenshot":       u"screenshot.png" if self._screenshot_path else None,
                "log":              _log,
            }

            # Verifica que la unidad/ruta raíz exista antes de guardar
            _drive = os.path.splitdrive(self._issues_dir)[0]
            if _drive and not os.path.exists(_drive + "\\"):
                raise Exception(
                    u"La unidad '{}' no esta disponible. "
                    u"Conecta la unidad o cambia ISSUES_DIR en el script.".format(_drive)
                )

            # ── Guardado inline (no usa funciones de módulo) ──────────────────
            issue_data["elementos"] = [
                {"id": e.ElementId, "categoria": e.Categoria, "tipo": e.Tipo}
                for e in list(self._elementos)
            ]
            # Guardar dentro de la subcarpeta del proyecto
            issue_dir   = os.path.join(self._project_issues_dir, issue_data["id"])
            issues_root = self._project_issues_dir

            import System.IO as _sio
            import System as _sys
            if not _sio.Directory.Exists(issues_root):
                _sio.Directory.CreateDirectory(issues_root)
            _sio.Directory.CreateDirectory(issue_dir)

            if self._screenshot_path and _sio.File.Exists(self._screenshot_path):
                _sio.File.Copy(
                    self._screenshot_path,
                    os.path.join(issue_dir, "screenshot.png"),
                    True
                )

            # Copiar archivos adjuntos a issue_dir/adjuntos/
            _adjuntos_guardados = []
            _adjuntos_lista = list(self._adjuntos_col)
            if _adjuntos_lista:
                _adj_dir = os.path.join(issue_dir, "adjuntos")
                _sio.Directory.CreateDirectory(_adj_dir)
                for _adj in _adjuntos_lista:
                    try:
                        if _sio.File.Exists(_adj.Ruta):
                            _dst = os.path.join(_adj_dir, _adj.Nombre)
                            _sio.File.Copy(_adj.Ruta, _dst, True)
                            _adjuntos_guardados.append(_adj.Nombre)
                    except Exception:
                        pass
            issue_data["adjuntos"] = _adjuntos_guardados

            _sio.File.WriteAllText(
                os.path.join(issue_dir, "issue.json"),
                json.dumps(issue_data, ensure_ascii=False, indent=2),
                _sys.Text.Encoding.UTF8
            )
            # ─────────────────────────────────────────────────────────────────

            # Notificacion por correo al asignado - MAILTO (compatible Outlook clasico y nuevo)
            _correo_enviado = False
            _correo_error = None
            try:
                _email_to = u""
                if isinstance(asignado, dict):
                    _email_to = asignado.get(u"email", u"").strip()

                if _email_to:
                    _asi_nom = (asignado.get(u"nombre", u"") if isinstance(asignado, dict) else u"")

                    _subject = u"Incidencia N" + str(_issue_number) + u" - " + (titulo or u"") + u" - " + (prioridad or u"")

                    _body_lines = [
                        u"Estimado/a " + _asi_nom + u",",
                        u"",
                        u"Se te ha asignado una nueva incidencia en el proyecto " + (proyecto or u"") + u".",
                        u"",
                        u"Incidencia N" + str(_issue_number),
                        u"",
                        titulo or u"",
                        descripcion or u"",
                        u"",
                        u"Por favor, revisar la aplicacion revisar incidencias para mas detalles.",
                    ]
                    _body_plain = u"\r\n".join(_body_lines)

                    _logo_uri, _logo_w, _logo_h = _load_logo_base64_bimissue()
                    _body_html = _build_email_html_bimissue(
                        _asi_nom, proyecto, _issue_number, titulo, descripcion, _logo_uri, _logo_w, _logo_h
                    )

                    _email_from = (reportado.get(u"email", u"") if isinstance(reportado, dict) else u"").strip()
                    if _send_mailto_bimissue(_email_to, _subject, _body_plain, _body_html, _email_from):
                        _correo_enviado = True
                    else:
                        _correo_error = u"mailto fallo"
                else:
                    _correo_error = u"Asignado a no tiene email configurado"

            except Exception as _ex:
                _correo_error = str(_ex)

            # Mensaje de estado con resultado del correo
            _estado_msg = u"Incidencia guardada en: {}".format(issue_dir)
            if _correo_enviado:
                _estado_msg += u" | Borrador de correo abierto para {} (revisa y envia)".format(_email_to)
            elif _correo_error:
                _estado_msg += u" | Correo no abierto: {}".format(_correo_error)
            self._set_estado(_estado_msg)
            self._win.Close()

        except Exception as ex:
            _ex_msg = str(ex)
            self._set_estado(u"ERROR: " + _ex_msg)
            try:
                open(u"C:\\Users\\jinun\\BIMISSUE_ERR.txt", "w").write(str(ex))
            except Exception:
                pass

    def _persona_from_combo(self, combo):
        sel = combo.SelectedItem
        if isinstance(sel, self._PersonaItem):
            return {"nombre": sel.Nombre, "email": sel.Email}
        text = (combo.Text or u"").strip()
        return {"nombre": text or u"—", "email": u""}

    def _get_disciplinas(self):
        mapping = [
            (u"Arquitectura", self._chk_arquitectura),
            (u"Estructura",   self._chk_estructura),
            (u"MEP",          self._chk_mep),
            (u"Civil",        self._chk_civil),
            (u"General",      self._chk_general),
        ]
        selected = [name for name, chk in mapping if chk and chk.IsChecked]
        return selected if selected else [u"General"]

    @staticmethod
    def _combo_tag(combo, default):
        try:
            item = combo.SelectedItem
            if item is None:
                return default
            tag = getattr(item, "Tag", None)
            if tag is not None:
                return str(tag)
            content = getattr(item, "Content", None)
            if content is not None:
                return str(content)
            return str(item) or default
        except Exception:
            return default

    def ShowDialog(self):
        self._win.ShowDialog()

    def _set_estado_ext(self, msg):
        self._set_estado(msg)


# ── Entry point ─────────────────────────────────────────────────────────────────
try:
    _active_view   = uidoc.ActiveView
    _initial_png   = None
    _capture_error = None

    try:
        _initial_png = _capture_view(_active_view)
    except Exception as _ex:
        _capture_error = u"Error al capturar: {}".format(str(_ex))

    w = BIMIssueWindow(initial_png=_initial_png, initial_view=_active_view)

    if _capture_error:
        w._set_estado(_capture_error)

    w.ShowDialog()

except Exception as ex:
    import traceback
    try:
        from pyrevit import forms as _pf
        _pf.alert(
            u"Error al iniciar BIMIssue:\n\n{}\n\n{}".format(
                str(ex), traceback.format_exc()),
            title=u"Error")
    except Exception:
        MessageBox.Show(
            u"{}\n\n{}".format(str(ex), traceback.format_exc()),
            u"Error — BIMIssue")
