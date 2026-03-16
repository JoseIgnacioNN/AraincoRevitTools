# -*- coding: ascii -*-
"""
VerIssues - Consulta todas las incidencias creadas en el servidor BIM.
"""

__title__ = "Ver\nIncidencias"
__author__ = "pyRevit"
__doc__    = "Consulta, filtra y visualiza todas las incidencias del servidor."

import os
import json

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Collections")

from System.Windows.Markup          import XamlReader
from System.Windows                 import Window, RoutedEventHandler, MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows.Controls        import SelectionChangedEventHandler
from System.Windows.Input           import Key, KeyEventHandler
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Media.Imaging   import BitmapImage
from System                         import Uri, UriKind
import System.IO  as sio
import System

ISSUES_DIR    = u"Y:\\00_SERVIDOR DE INCIDENCIAS"
PERSONAS_FILE = os.path.join(ISSUES_DIR, "personas.json")

# Ruta al logo: busca primero en esta carpeta, luego en BIMIssue (fallback)
try:
    _THIS_DIR  = os.path.dirname(os.path.abspath(__file__))
    _PANEL_DIR = os.path.dirname(_THIS_DIR)
    _LOGO_PATHS = [
        os.path.join(_THIS_DIR,  "logo.png"),
        os.path.join(_PANEL_DIR, "01_BIMIssue.pushbutton", "logo.png"),
    ]
except Exception:
    _LOGO_PATHS = []

# Revit API - disponible en scope de modulo al cargar el script
try:
    from Autodesk.Revit.DB import (
        FilteredElementCollector, View3D, ViewFamilyType,
        ViewFamily, XYZ, ViewOrientation3D, Transaction,
        ModelPathUtils
    )
    _uidoc = __revit__.ActiveUIDocument
    _doc   = _uidoc.Document
except Exception:
    FilteredElementCollector = None
    View3D = None
    ViewFamilyType = None
    ViewFamily = None
    XYZ = None
    ViewOrientation3D = None
    Transaction = None
    ModelPathUtils = None
    _uidoc = None
    _doc   = None


def _get_project_folder_name(document):
    """
    Devuelve un nombre de carpeta limpio para el proyecto activo.

    Mismo formato que Crear Incidencia para que el filtro por proyecto coincida:
        "DIRECTORIO QUE CONTIENE EL MODELO CENTRAL" _ "NOMBRE DEL MODELO CENTRAL"

    Lógica:
    1. Del archivo actual se obtiene el modelo central (si es workshared) o se usa
       la ruta del archivo actual (si es independiente).
    2. Del modelo central se obtiene: el nombre del directorio que lo contiene y
       el nombre del archivo del modelo central (sin extensión).
    Los caracteres inválidos para nombres de carpeta se reemplazan por '_'.
    """
    import re
    folder_name = u""
    try:
        model_title = None
        try:
            model_title = (document.Title or u"").strip() if document else None
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

        if document and ModelPathUtils and document.IsWorkshared:
            central_path = document.GetWorksharingCentralModelPath()
            user_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(central_path)
            folder_name = _from_path(user_path, model_title)
        if not folder_name and document and document.PathName:
            # Archivo independiente: la ruta del archivo actual es la del "modelo central"
            folder_name = _from_path(document.PathName, model_title)
        if not folder_name:
            base = (model_title or u"SinNombre").strip() or u"SinNombre"
            folder_name = u"{0}_{0}".format(base)
    except Exception:
        try:
            base = (document.Title or u"SinNombre").strip() or u"SinNombre" if document else u"SinNombre"
            folder_name = u"{0}_{0}".format(base)
        except Exception:
            folder_name = u"SinNombre_SinNombre"
    # Eliminar caracteres inválidos en nombres de carpeta Windows
    folder_name = re.sub(r'[<>:"/\\|?*]', u"_", folder_name).strip(u". ")
    return folder_name if folder_name else u"SinNombre_SinNombre"


EM_DASH = u"\u2014"
N_ORDINAL = u"\u00ba"


# -- Personas (para combo asignado a) -------------------------------------------
class PersonaItem(object):
    def __init__(self, nombre, email=u""):
        self.Nombre = nombre
        self.Email  = email

    def __str__(self):
        return self.Nombre

    def ToString(self):
        return self.Nombre


def _load_personas(personas_file):
    """Lee personas.json y devuelve lista de PersonaItem."""
    personas = []
    if os.path.exists(personas_file):
        try:
            raw  = sio.File.ReadAllText(personas_file, System.Text.Encoding.UTF8)
            data = json.loads(raw)
            for p in data:
                personas.append(PersonaItem(p.get("nombre", u""), p.get("email", u"")))
        except Exception:
            pass
    return personas


# -- Modelos de datos ----------------------------------------------------------
class DetAdjuntoItem(object):
    """Adjunto mostrado en el panel de detalle de Ver Incidencias."""
    def __init__(self, nombre, ruta_completa):
        self.Nombre = nombre
        self.Ruta   = ruta_completa

    def ToString(self):
        return self.Nombre


class IssueItem(object):
    def __init__(self, data, issue_dir):
        import os
        self.Numero       = u"{}".format(data.get("numero", u"?"))
        self.Proyecto     = data.get("proyecto_carpeta", data.get("proyecto", EM_DASH))
        self.Titulo       = data.get("titulo", EM_DASH)
        self.Prioridad    = data.get("prioridad", EM_DASH)
        self.Estado       = data.get("estado", EM_DASH)
        _rep = data.get("reportado_por", {})
        self.ReportadoPor = _rep.get("nombre", EM_DASH) if isinstance(_rep, dict) else str(_rep)
        _asi = data.get("asignado_a", {})
        self.AsignadoA    = _asi.get("nombre", EM_DASH) if isinstance(_asi, dict) else str(_asi)
        _fecha = data.get("fecha", u"")
        _fiso  = _fecha[:10] if len(_fecha) >= 10 else _fecha
        try:
            _parts = _fiso.split(u"-")
            self.Fecha = u"{}/{}/{}".format(_parts[2], _parts[1], _parts[0])
        except Exception:
            self.Fecha = _fiso
        self.IssueDir     = issue_dir
        _ss = data.get("screenshot")
        self.ScreenshotPath = os.path.join(issue_dir, "screenshot.png") if _ss else None
        self._data = data

    def ToString(self):
        return self.Titulo


# -- Carga de incidencias ------------------------------------------------------
def _load_all_issues(issues_dir):
    import os
    import json
    import System.IO  as _sio
    import System     as _sys
    issues = []
    if not os.path.isdir(issues_dir):
        return issues
    for project_name in sorted(os.listdir(issues_dir)):
        project_path = os.path.join(issues_dir, project_name)
        if not os.path.isdir(project_path):
            continue
        if "." in project_name:
            continue
        for issue_name in sorted(os.listdir(project_path), reverse=False):
            if not issue_name.startswith("ISSUE_"):
                continue
            issue_path = os.path.join(project_path, issue_name)
            if not os.path.isdir(issue_path):
                continue
            json_path = os.path.join(issue_path, "issue.json")
            if not os.path.exists(json_path):
                continue
            try:
                raw  = _sio.File.ReadAllText(json_path, _sys.Text.Encoding.UTF8)
                data = json.loads(raw)
                # Fallback: proyecto desde carpeta si no está en JSON
                if not data.get("proyecto_carpeta") and not data.get("proyecto"):
                    data = dict(data)
                    data["proyecto_carpeta"] = project_name
                issues.append(IssueItem(data, issue_path))
            except Exception:
                pass
    return issues


# -- XAML ----------------------------------------------------------------------
XAML = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Consulta de Incidencias"
    Height="1070" Width="1280"
    MinHeight="600" MinWidth="900"
    WindowStartupLocation="CenterScreen"
    Background="#0A1C26"
    FontFamily="Segoe UI"
    ResizeMode="CanResize">

  <Window.Resources>
    <Style x:Key="FilterLbl" TargetType="TextBlock">
      <Setter Property="Foreground"        Value="#4A8BA6"/>
      <Setter Property="FontSize"          Value="10"/>
      <Setter Property="FontWeight"        Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin"            Value="16,0,6,0"/>
    </Style>
    <Style x:Key="FilterCombo" TargetType="ComboBox">
      <Setter Property="Background"      Value="#0D2234"/>
      <Setter Property="Foreground"      Value="#F0F8FC"/>
      <Setter Property="BorderBrush"     Value="#1A3D52"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Height"          Value="30"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid TextElement.Foreground="{TemplateBinding Foreground}">
              <Border x:Name="Bd"
                      Background="{TemplateBinding Background}"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}"
                      CornerRadius="5"/>
              <ToggleButton IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay,
                            RelativeSource={RelativeSource TemplatedParent}}"
                            Focusable="False">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border Background="Transparent"/>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter IsHitTestVisible="False"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                Margin="10,0,28,0" VerticalAlignment="Center"/>
              <TextBlock Text="&#9660;" FontSize="8" Foreground="#5BB8D4"
                         HorizontalAlignment="Right" VerticalAlignment="Center"
                         Margin="0,0,10,0" IsHitTestVisible="False"/>
              <Popup x:Name="PART_Popup" IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False"
                     PopupAnimation="Fade" Placement="Bottom">
                <Border Background="#0D2234" BorderBrush="#1A3D52" BorderThickness="1"
                        MinWidth="{Binding ActualWidth,
                        RelativeSource={RelativeSource TemplatedParent}}">
                  <ScrollViewer MaxHeight="200" VerticalScrollBarVisibility="Auto">
                    <ItemsPresenter/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver"    Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="FilterComboItem" TargetType="ComboBoxItem">
      <Setter Property="Background" Value="#0D2234"/>
      <Setter Property="Foreground" Value="#F0F8FC"/>
      <Setter Property="Padding"    Value="10,6"/>
      <Setter Property="FontSize"   Value="12"/>
      <Style.Triggers>
        <Trigger Property="IsHighlighted" Value="True">
          <Setter Property="Background" Value="#1A4F6A"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="ActionBtn" TargetType="Button">
      <Setter Property="Background"      Value="#1A3D52"/>
      <Setter Property="Foreground"      Value="#C8E4EF"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontSize"        Value="11"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="Padding"         Value="14,0"/>
      <Setter Property="Height"          Value="30"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                    CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#235E7D"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#1A4F6A"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="DetLbl" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#4A8BA6"/>
      <Setter Property="FontSize"   Value="10"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin"     Value="0,10,0,2"/>
    </Style>
    <Style x:Key="DetVal" TargetType="TextBlock">
      <Setter Property="Foreground"   Value="#C8E4EF"/>
      <Setter Property="FontSize"     Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" Background="#0F2535" Padding="20,10">
      <DockPanel>
        <!-- Oculto (ruta servidor) -->
        <TextBlock DockPanel.Dock="Right" x:Name="TxtServidor"
                   Foreground="#2A5570" FontSize="10" VerticalAlignment="Center"
                   Visibility="Collapsed"/>

        <!-- Titulo + badge -->
        <StackPanel DockPanel.Dock="Left" Orientation="Horizontal" VerticalAlignment="Center">
          <Image x:Name="ImgLogo" Width="80" Height="80" Margin="0,0,16,0"
                 Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality"/>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="CONSULTA DE INCIDENCIAS" FontSize="18" FontWeight="Bold"
                       Foreground="#C8E4EF"/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
              <Border Background="#1A4F6A" CornerRadius="8" Padding="8,2" Margin="0,0,8,0">
                <TextBlock x:Name="TxtTotal" Foreground="#5BB8D4" FontSize="11"
                           FontWeight="Bold" Text="0 issues"/>
              </Border>
            </StackPanel>
          </StackPanel>
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- Barra de filtros -->
    <Border Grid.Row="1" Background="#091820" BorderBrush="#1A3D52"
            BorderThickness="0,0,0,1" Padding="4,8">
      <StackPanel Orientation="Horizontal">
        <TextBlock Text="Proyecto" Style="{StaticResource FilterLbl}" Margin="12,0,6,0"/>
        <ComboBox x:Name="CmbFiltroProyecto" Width="200"
                  Style="{StaticResource FilterCombo}"/>

        <TextBlock Text="Estado" Style="{StaticResource FilterLbl}"/>
        <ComboBox x:Name="CmbFiltroEstado" Width="130"
                  Style="{StaticResource FilterCombo}">
          <ComboBox.ItemContainerStyle>
            <Style TargetType="ComboBoxItem" BasedOn="{StaticResource FilterComboItem}"/>
          </ComboBox.ItemContainerStyle>
          <ComboBoxItem Content="Todos"       IsSelected="True"/>
          <ComboBoxItem Content="Abierto"/>
          <ComboBoxItem Content="En revision"/>
          <ComboBoxItem Content="Resuelto"/>
        </ComboBox>

        <TextBlock Text="Prioridad" Style="{StaticResource FilterLbl}"/>
        <ComboBox x:Name="CmbFiltroPrioridad" Width="120"
                  Style="{StaticResource FilterCombo}">
          <ComboBox.ItemContainerStyle>
            <Style TargetType="ComboBoxItem" BasedOn="{StaticResource FilterComboItem}"/>
          </ComboBox.ItemContainerStyle>
          <ComboBoxItem Content="Todos"   IsSelected="True"/>
          <ComboBoxItem Content="Critica"/>
          <ComboBoxItem Content="Alta"/>
          <ComboBoxItem Content="Media"/>
          <ComboBoxItem Content="Baja"/>
        </ComboBox>

        <Button x:Name="BtnLimpiarFiltros" Content="Limpiar filtros"
                Style="{StaticResource ActionBtn}" Margin="16,0,0,0"/>
        <Button x:Name="BtnRefresh" Content="Actualizar"
                Style="{StaticResource ActionBtn}" Margin="8,0,0,0"
                Background="#1A4F6A"/>
        <Button x:Name="BtnDashboard" Content="Dashboard"
                Style="{StaticResource ActionBtn}" Margin="8,0,0,0"
                Background="#2A3F5A"/>
        <Button x:Name="BtnManual" Content="Manual"
                Style="{StaticResource ActionBtn}" Margin="8,0,0,0"
                Background="#2A5C3D" ToolTip="Abrir manual de usuario"/>
      </StackPanel>
    </Border>

    <!-- Contenido principal -->
    <Grid Grid.Row="2">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="360"/>
      </Grid.ColumnDefinitions>

      <!-- Placeholder: sin incidencias -->
      <Border Grid.Column="0" x:Name="GridPlaceholder"
              Background="#081520" Visibility="Collapsed" Panel.ZIndex="1">
        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
          <TextBlock Text="Sin incidencias" FontSize="20" FontWeight="SemiBold"
                     Foreground="#1A3D52" HorizontalAlignment="Center"/>
          <TextBlock FontSize="12" Foreground="#2A5570"
                     HorizontalAlignment="Center" Margin="0,10,0,0" TextWrapping="Wrap"
                     MaxWidth="320" TextAlignment="Center"
                     Text="No se encontraron incidencias en el servidor.&#10;Usa BIM Issue para crear la primera."/>
        </StackPanel>
      </Border>

      <!-- Lista de incidencias -->
      <DataGrid Grid.Column="0" x:Name="GridIssues"
                AutoGenerateColumns="False" IsReadOnly="True"
                SelectionMode="Single" CanUserAddRows="False"
                CanUserReorderColumns="False" CanUserResizeRows="False"
                Background="#081520" Foreground="#C8E4EF"
                BorderThickness="0"
                GridLinesVisibility="Horizontal"
                HorizontalGridLinesBrush="#132C3D"
                RowBackground="#081520"
                AlternatingRowBackground="#0B1F2E"
                ColumnHeaderHeight="32" RowHeight="36" FontSize="12"
                HorizontalScrollBarVisibility="Auto"
                VerticalScrollBarVisibility="Auto">
        <DataGrid.Columns>
          <DataGridTextColumn Header="N&#186;" Binding="{Binding Numero}"  Width="50"/>
          <DataGridTextColumn Header="Titulo"   Binding="{Binding Titulo}"   Width="*"/>

          <!-- Prioridad con badge de color -->
          <DataGridTemplateColumn Header="Prioridad" Width="90" CanUserResize="False">
            <DataGridTemplateColumn.CellTemplate>
              <DataTemplate>
                <Border CornerRadius="4" Padding="7,2" HorizontalAlignment="Center" Margin="0,4">
                  <Border.Style>
                    <Style TargetType="Border">
                      <Setter Property="Background" Value="#1A3D52"/>
                      <Style.Triggers>
                        <DataTrigger Binding="{Binding Prioridad}" Value="Critica">
                          <Setter Property="Background" Value="#6B1A1A"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Prioridad}" Value="Alta">
                          <Setter Property="Background" Value="#7A3B0A"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Prioridad}" Value="Media">
                          <Setter Property="Background" Value="#5B4E08"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Prioridad}" Value="Baja">
                          <Setter Property="Background" Value="#1A4F2A"/>
                        </DataTrigger>
                      </Style.Triggers>
                    </Style>
                  </Border.Style>
                  <TextBlock Text="{Binding Prioridad}" FontSize="11" FontWeight="SemiBold"
                             HorizontalAlignment="Center">
                    <TextBlock.Style>
                      <Style TargetType="TextBlock">
                        <Setter Property="Foreground" Value="#C8E4EF"/>
                        <Style.Triggers>
                          <DataTrigger Binding="{Binding Prioridad}" Value="Critica">
                            <Setter Property="Foreground" Value="#FFB4B4"/>
                          </DataTrigger>
                          <DataTrigger Binding="{Binding Prioridad}" Value="Alta">
                            <Setter Property="Foreground" Value="#FFB76B"/>
                          </DataTrigger>
                          <DataTrigger Binding="{Binding Prioridad}" Value="Media">
                            <Setter Property="Foreground" Value="#FFE47A"/>
                          </DataTrigger>
                          <DataTrigger Binding="{Binding Prioridad}" Value="Baja">
                            <Setter Property="Foreground" Value="#6BCC8E"/>
                          </DataTrigger>
                        </Style.Triggers>
                      </Style>
                    </TextBlock.Style>
                  </TextBlock>
                </Border>
              </DataTemplate>
            </DataGridTemplateColumn.CellTemplate>
          </DataGridTemplateColumn>

          <!-- Estado con badge de color -->
          <DataGridTemplateColumn Header="Estado" Width="100" CanUserResize="False">
            <DataGridTemplateColumn.CellTemplate>
              <DataTemplate>
                <Border CornerRadius="4" Padding="7,2" HorizontalAlignment="Center" Margin="0,4">
                  <Border.Style>
                    <Style TargetType="Border">
                      <Setter Property="Background" Value="#1A3D52"/>
                      <Style.Triggers>
                        <DataTrigger Binding="{Binding Estado}" Value="Abierto">
                          <Setter Property="Background" Value="#1A3D70"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Estado}" Value="En revision">
                          <Setter Property="Background" Value="#3D1A6B"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Estado}" Value="Resuelto">
                          <Setter Property="Background" Value="#1A4F2A"/>
                        </DataTrigger>
                      </Style.Triggers>
                    </Style>
                  </Border.Style>
                  <TextBlock Text="{Binding Estado}" FontSize="11" FontWeight="SemiBold"
                             HorizontalAlignment="Center">
                    <TextBlock.Style>
                      <Style TargetType="TextBlock">
                        <Setter Property="Foreground" Value="#C8E4EF"/>
                        <Style.Triggers>
                          <DataTrigger Binding="{Binding Estado}" Value="Abierto">
                            <Setter Property="Foreground" Value="#7EB8FF"/>
                          </DataTrigger>
                          <DataTrigger Binding="{Binding Estado}" Value="En revision">
                            <Setter Property="Foreground" Value="#C97EFF"/>
                          </DataTrigger>
                          <DataTrigger Binding="{Binding Estado}" Value="Resuelto">
                            <Setter Property="Foreground" Value="#6BCC8E"/>
                          </DataTrigger>
                        </Style.Triggers>
                      </Style>
                    </TextBlock.Style>
                  </TextBlock>
                </Border>
              </DataTemplate>
            </DataGridTemplateColumn.CellTemplate>
          </DataGridTemplateColumn>

          <DataGridTextColumn Header="Asignado a"      Binding="{Binding AsignadoA}" Width="130"/>
          <DataGridTextColumn Header="Fecha de Creaci&#243;n" Binding="{Binding Fecha}" Width="120"/>
        </DataGrid.Columns>
        <DataGrid.ColumnHeaderStyle>
          <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background"      Value="#0F2535"/>
            <Setter Property="Foreground"      Value="#5BB8D4"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="10,0"/>
            <Setter Property="BorderBrush"     Value="#1A3D52"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
          </Style>
        </DataGrid.ColumnHeaderStyle>
        <DataGrid.CellStyle>
          <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness"   Value="0"/>
            <Setter Property="Padding"           Value="10,0"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Style.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="Foreground" Value="#F0F8FC"/>
              </Trigger>
            </Style.Triggers>
          </Style>
        </DataGrid.CellStyle>
        <DataGrid.RowStyle>
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
            </Style.Triggers>
          </Style>
        </DataGrid.RowStyle>
      </DataGrid>

      <!-- Panel de detalle -->
      <Border Grid.Column="1" Background="#091820" BorderBrush="#1A3D52"
              BorderThickness="1,0,0,0">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel Margin="16,14">

            <!-- Placeholder -->
            <Border x:Name="DetPlaceholder" Background="#0F2535" CornerRadius="8"
                    Height="160" Margin="0,60,0,0">
              <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
                <TextBlock Text="[ ]" FontSize="30" HorizontalAlignment="Center"
                           Foreground="#1A3D52"/>
                <TextBlock Text="Selecciona una incidencia" FontSize="12"
                           Foreground="#2A5570" HorizontalAlignment="Center"
                           Margin="0,8,0,0"/>
              </StackPanel>
            </Border>

            <!-- Contenido de detalle (visible al seleccionar) -->
            <StackPanel x:Name="DetContenido" Visibility="Collapsed">

              <!-- Screenshot -->
              <Border x:Name="DetImgPlaceholder" Background="#0F2535" CornerRadius="6"
                      Height="180" Margin="0,0,0,10">
                <TextBlock Text="Sin captura" FontSize="11" Foreground="#2A5570"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <Image x:Name="ImgDetail" Height="180" Margin="0,0,0,10"
                     Stretch="UniformToFill" StretchDirection="DownOnly"
                     Visibility="Collapsed"/>

              <!-- Cabecera del issue -->
              <Border Background="#0F2535" CornerRadius="6" Padding="12,8" Margin="0,0,0,4">
                <StackPanel>
                  <StackPanel Orientation="Horizontal">
                    <TextBlock x:Name="DetBadgeNum"
                               Foreground="#5BB8D4" FontSize="11" FontWeight="Bold"
                               Margin="0,0,8,0" VerticalAlignment="Center"/>
                    <TextBlock x:Name="DetBadgeProyecto"
                               Foreground="#4A8BA6" FontSize="10"
                               VerticalAlignment="Center"/>
                  </StackPanel>
                  <TextBlock x:Name="DetTitulo" Foreground="#F0F8FC" FontSize="14"
                             FontWeight="Bold" TextWrapping="Wrap" Margin="0,6,0,0"/>
                </StackPanel>
              </Border>

              <!-- Badges prioridad + estado -->
              <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
                <Border x:Name="DetBdPrioridad" CornerRadius="4" Padding="10,3"
                        Margin="0,0,6,0">
                  <TextBlock x:Name="DetPrioridad" FontSize="11" FontWeight="SemiBold"/>
                </Border>
                <Border x:Name="DetBdEstado" CornerRadius="4" Padding="10,3">
                  <TextBlock x:Name="DetEstado" FontSize="11" FontWeight="SemiBold"/>
                </Border>
              </StackPanel>

              <!-- Cambiar prioridad -->
              <Border Background="#0D1E2A" CornerRadius="6" Padding="10,8"
                      Margin="0,8,0,0">
                <StackPanel>
                  <TextBlock Text="CAMBIAR PRIORIDAD" FontSize="10" FontWeight="SemiBold"
                             Foreground="#4A8BA6" Margin="0,0,0,6"/>
                  <StackPanel Orientation="Horizontal">
                    <ComboBox x:Name="CmbDetPrioridad" Width="150" Height="28"
                              IsEnabled="False" Style="{StaticResource FilterCombo}">
                      <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem"
                               BasedOn="{StaticResource FilterComboItem}"/>
                      </ComboBox.ItemContainerStyle>
                      <ComboBoxItem Content="Critica"/>
                      <ComboBoxItem Content="Alta"/>
                      <ComboBoxItem Content="Media"/>
                      <ComboBoxItem Content="Baja"/>
                    </ComboBox>
                    <TextBlock x:Name="TxtPrioridadConfirm"
                               Text="  Prioridad guardada" Foreground="#6BCC8E"
                               FontSize="10" FontWeight="SemiBold"
                               VerticalAlignment="Center"
                               Visibility="Collapsed"/>
                  </StackPanel>
                </StackPanel>
              </Border>

              <!-- Cambiar estado -->
              <Border Background="#0D1E2A" CornerRadius="6" Padding="10,8"
                      Margin="0,8,0,0">
                <StackPanel>
                  <TextBlock Text="CAMBIAR ESTADO" FontSize="10" FontWeight="SemiBold"
                             Foreground="#4A8BA6" Margin="0,0,0,6"/>
                  <StackPanel Orientation="Horizontal">
                    <ComboBox x:Name="CmbDetEstado" Width="150" Height="28"
                              IsEnabled="False" Style="{StaticResource FilterCombo}">
                      <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem"
                               BasedOn="{StaticResource FilterComboItem}"/>
                      </ComboBox.ItemContainerStyle>
                      <ComboBoxItem Content="Abierto"/>
                      <ComboBoxItem Content="En revision"/>
                      <ComboBoxItem Content="Resuelto"/>
                    </ComboBox>
                    <TextBlock x:Name="TxtEstadoConfirm"
                               Text="  Estado guardado" Foreground="#6BCC8E"
                               FontSize="10" FontWeight="SemiBold"
                               VerticalAlignment="Center"
                               Visibility="Collapsed"/>
                  </StackPanel>
                </StackPanel>
              </Border>

              <!-- Cambiar asignado a -->
              <Border Background="#0D1E2A" CornerRadius="6" Padding="10,8"
                      Margin="0,8,0,0">
                <StackPanel>
                  <TextBlock Text="CAMBIAR ASIGNADO A" FontSize="10" FontWeight="SemiBold"
                             Foreground="#4A8BA6" Margin="0,0,0,6"/>
                  <StackPanel Orientation="Horizontal">
                    <ComboBox x:Name="CmbDetAsignado" Width="200" Height="28"
                              IsEnabled="False" Style="{StaticResource FilterCombo}"
                              DisplayMemberPath="Nombre">
                      <ComboBox.ItemContainerStyle>
                        <Style TargetType="ComboBoxItem"
                               BasedOn="{StaticResource FilterComboItem}"/>
                      </ComboBox.ItemContainerStyle>
                    </ComboBox>
                    <TextBlock x:Name="TxtAsignadoConfirm"
                               Text="  Asignado guardado" Foreground="#6BCC8E"
                               FontSize="10" FontWeight="SemiBold"
                               VerticalAlignment="Center"
                               Visibility="Collapsed"/>
                  </StackPanel>
                </StackPanel>
              </Border>

              <!-- Campos -->
              <TextBlock Style="{StaticResource DetLbl}" Text="DESCRIPCION"/>
              <TextBlock x:Name="DetDescripcion" Style="{StaticResource DetVal}"/>

              <TextBlock Style="{StaticResource DetLbl}" Text="DISCIPLINAS"/>
              <TextBlock x:Name="DetDisciplinas" Style="{StaticResource DetVal}"/>

              <Grid Margin="0,6,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Margin="0,0,8,0">
                  <TextBlock Style="{StaticResource DetLbl}" Text="REPORTADO POR"
                             Margin="0,4,0,2"/>
                  <TextBlock x:Name="DetReportadoPor" Style="{StaticResource DetVal}"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <TextBlock Style="{StaticResource DetLbl}" Text="ASIGNADO A"
                             Margin="0,4,0,2"/>
                  <TextBlock x:Name="DetAsignadoA" Style="{StaticResource DetVal}"/>
                </StackPanel>
              </Grid>

              <Border Height="1" Background="#1A3D52" Margin="0,12,0,0"/>

              <Grid Margin="0,4,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                  <TextBlock Style="{StaticResource DetLbl}" Text="FECHA DE CREACI&#211;N" Margin="0,4,0,2"/>
                  <TextBlock x:Name="DetFecha" Style="{StaticResource DetVal}"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <TextBlock Style="{StaticResource DetLbl}" Text="AUTOR" Margin="0,4,0,2"/>
                  <TextBlock x:Name="DetAutor" Style="{StaticResource DetVal}"/>
                </StackPanel>
              </Grid>

              <TextBlock Style="{StaticResource DetLbl}" Text="ELEMENTOS VINCULADOS"/>
              <TextBlock x:Name="DetElementos" Style="{StaticResource DetVal}"
                         Foreground="#4A8BA6"/>

              <TextBlock Style="{StaticResource DetLbl}" Text="VISTA / CAMARA"/>
              <TextBlock x:Name="DetVista" Style="{StaticResource DetVal}"
                         FontSize="11" Foreground="#4A8BA6"/>

              <!-- Adjuntos -->
              <TextBlock Style="{StaticResource DetLbl}" Text="ARCHIVOS ADJUNTOS"/>
              <Border x:Name="DetAdjuntosBox" Background="#0D1E2A" CornerRadius="5"
                      Padding="8,6" Margin="0,0,0,2">
                <StackPanel>
                  <TextBlock x:Name="DetSinAdjuntos" Text="Sin archivos adjuntos"
                             FontSize="11" Foreground="#2A5570"/>
                  <ListBox x:Name="LstDetAdjuntos" Background="Transparent"
                           BorderThickness="0" MaxHeight="90"
                           DisplayMemberPath="Nombre" Visibility="Collapsed"
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
                  <Button x:Name="BtnAbrirAdjunto" Content="Abrir archivo"
                          Style="{StaticResource ActionBtn}"
                          Background="#1A3D52" Height="28" FontSize="11"
                          IsEnabled="False" Margin="0,8,0,0"/>
                </StackPanel>
              </Border>

              <!-- Boton ir al punto de vista -->
              <Border Height="1" Background="#1A3D52" Margin="0,14,0,0"/>
              <Button x:Name="BtnIrViewpoint"
                      Content="Ir al punto de vista"
                      Style="{StaticResource ActionBtn}"
                      Background="#1A4F6A"
                      Height="38" FontSize="12"
                      IsEnabled="False"
                      Margin="0,12,0,6"/>
            </StackPanel>
          </StackPanel>
        </ScrollViewer>
      </Border>
    </Grid>

    <!-- Footer -->
    <Border Grid.Row="3" Background="#0A1C26" BorderBrush="#1A3D52"
            BorderThickness="0,1,0,0" Padding="16,8">
      <TextBlock x:Name="TxtStatus" FontSize="11" Foreground="#4A8BA6"
                 Text="Listo."/>
    </Border>
  </Grid>
</Window>
"""


# -- Dashboard de Incidencias --------------------------------------------------
class DashboardWindow(object):

    # ---- paletas de color ----------------------------------------------------
    _PRIOR_COLORS = {
        u"Critica": u"#8B2020",
        u"Alta":    u"#A05010",
        u"Media":   u"#7A6A10",
        u"Baja":    u"#1A6B35",
    }
    _PALETTE = [
        u"#1A4F6A", u"#4A1A6A", u"#1A6A4A", u"#6A4A1A",
        u"#6A1A3A", u"#2A3A6A", u"#3A5A2A", u"#5A3A2A",
    ]

    # ---- XAML ----------------------------------------------------------------
    _XAML = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Dashboard de Incidencias"
    Height="700" Width="860"
    MinHeight="500" MinWidth="700"
    WindowStartupLocation="CenterScreen"
    Background="#0A1C26"
    FontFamily="Segoe UI"
    ResizeMode="CanResize">

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" Background="#0F2535" Padding="20,12">
      <DockPanel>
        <TextBlock DockPanel.Dock="Left"
                   Text="DASHBOARD DE INCIDENCIAS"
                   FontSize="18" FontWeight="Bold"
                   Foreground="#C8E4EF" VerticalAlignment="Center"/>
        <TextBlock x:Name="DashTxtTotal" DockPanel.Dock="Right"
                   Foreground="#5BB8D4" FontSize="13" FontWeight="Bold"
                   VerticalAlignment="Center" HorizontalAlignment="Right"/>
      </DockPanel>
    </Border>

    <!-- Grid 2x2 de graficos -->
    <Grid Grid.Row="1" Margin="14">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Grid.RowDefinitions>
        <RowDefinition Height="*"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <!-- [0,0] Estado (dona) -->
      <Border Grid.Row="0" Grid.Column="0"
              Background="#0F2535" CornerRadius="8"
              Margin="0,0,7,7" Padding="16,12">
        <StackPanel>
          <TextBlock Text="ESTADO" Foreground="#4A8BA6"
                     FontSize="10" FontWeight="SemiBold" Margin="0,0,0,14"/>
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
            <!-- Dona -->
            <Canvas Width="130" Height="130">
              <Ellipse Width="130" Height="130" Fill="#152D40"/>
              <Path x:Name="DashPathAbierto"/>
              <Path x:Name="DashPathRevision"/>
              <Path x:Name="DashPathResuelto"/>
              <Ellipse Width="56" Height="56"
                       Canvas.Left="37" Canvas.Top="37" Fill="#0F2535"/>
            </Canvas>
            <!-- Leyenda -->
            <StackPanel VerticalAlignment="Center" Margin="20,0,0,0">
              <StackPanel Orientation="Horizontal" Margin="0,6">
                <Border Width="12" Height="12" Background="#1A3D70"
                        CornerRadius="3" Margin="0,0,8,0" VerticalAlignment="Center"/>
                <TextBlock Text="Abierto" Foreground="#7EB8FF"
                           FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <TextBlock x:Name="DashLblAbierto" Foreground="#7EB8FF"
                           FontSize="16" FontWeight="Bold"
                           VerticalAlignment="Center" Text="0"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,6">
                <Border Width="12" Height="12" Background="#3D1A6B"
                        CornerRadius="3" Margin="0,0,8,0" VerticalAlignment="Center"/>
                <TextBlock Text="En revision" Foreground="#C97EFF"
                           FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <TextBlock x:Name="DashLblRevision" Foreground="#C97EFF"
                           FontSize="16" FontWeight="Bold"
                           VerticalAlignment="Center" Text="0"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,6">
                <Border Width="12" Height="12" Background="#1A4F2A"
                        CornerRadius="3" Margin="0,0,8,0" VerticalAlignment="Center"/>
                <TextBlock Text="Resuelto" Foreground="#6BCC8E"
                           FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <TextBlock x:Name="DashLblResuelto" Foreground="#6BCC8E"
                           FontSize="16" FontWeight="Bold"
                           VerticalAlignment="Center" Text="0"/>
              </StackPanel>
            </StackPanel>
          </StackPanel>
        </StackPanel>
      </Border>

      <!-- [0,1] Prioridad (barras) -->
      <Border Grid.Row="0" Grid.Column="1"
              Background="#0F2535" CornerRadius="8"
              Margin="7,0,0,7" Padding="16,12">
        <StackPanel>
          <TextBlock Text="PRIORIDAD" Foreground="#4A8BA6"
                     FontSize="10" FontWeight="SemiBold" Margin="0,0,0,14"/>
          <Canvas x:Name="DashCanvasPrioridad" Width="340" Height="140"/>
        </StackPanel>
      </Border>

      <!-- [1,0] Proyecto (barras) -->
      <Border Grid.Row="1" Grid.Column="0"
              Background="#0F2535" CornerRadius="8"
              Margin="0,7,7,0" Padding="16,12">
        <StackPanel>
          <TextBlock Text="PROYECTO" Foreground="#4A8BA6"
                     FontSize="10" FontWeight="SemiBold" Margin="0,0,0,14"/>
          <Canvas x:Name="DashCanvasProyecto" Width="340" Height="200"/>
        </StackPanel>
      </Border>

      <!-- [1,1] Asignado a (barras) -->
      <Border Grid.Row="1" Grid.Column="1"
              Background="#0F2535" CornerRadius="8"
              Margin="7,7,0,0" Padding="16,12">
        <StackPanel>
          <TextBlock Text="ASIGNADO A" Foreground="#4A8BA6"
                     FontSize="10" FontWeight="SemiBold" Margin="0,0,0,14"/>
          <Canvas x:Name="DashCanvasAsignado" Width="340" Height="200"/>
        </StackPanel>
      </Border>

    </Grid>
  </Grid>
</Window>
"""

    # ---- constructor ---------------------------------------------------------
    def __init__(self, issues_list):
        self._issues = issues_list

        win = XamlReader.Parse(self._XAML)
        self._win = win

        self._txt_total         = win.FindName("DashTxtTotal")
        self._path_ab           = win.FindName("DashPathAbierto")
        self._path_rev          = win.FindName("DashPathRevision")
        self._path_res          = win.FindName("DashPathResuelto")
        self._lbl_ab            = win.FindName("DashLblAbierto")
        self._lbl_rev           = win.FindName("DashLblRevision")
        self._lbl_res           = win.FindName("DashLblResuelto")
        self._cv_prioridad      = win.FindName("DashCanvasPrioridad")
        self._cv_proyecto       = win.FindName("DashCanvasProyecto")
        self._cv_asignado       = win.FindName("DashCanvasAsignado")

        win.KeyDown += KeyEventHandler(self._on_key_down)
        win.Loaded  += RoutedEventHandler(self._on_loaded)

    # ---- carga inicial -------------------------------------------------------
    def _on_loaded(self, sender, args):
        total = len(self._issues)
        _s = u"s" if total != 1 else u""
        self._txt_total.Text = u"{} incidencia{}".format(total, _s)

        self._draw_donut()

        self._draw_bars(
            self._cv_prioridad,
            self._count_ordered(
                u"Prioridad",
                [u"Critica", u"Alta", u"Media", u"Baja"]
            ),
            lambda lbl: self._PRIOR_COLORS.get(lbl, u"#1A3D52"),
            max_bars=4
        )
        self._draw_bars(
            self._cv_proyecto,
            self._count_top(u"Proyecto"),
            lambda lbl: self._PALETTE[abs(hash(lbl)) % len(self._PALETTE)]
        )
        self._draw_bars(
            self._cv_asignado,
            self._count_top(u"AsignadoA"),
            lambda lbl: self._PALETTE[(abs(hash(lbl)) + 3) % len(self._PALETTE)]
        )

    # ---- agregacion de datos -------------------------------------------------
    def _count_ordered(self, attr, keys):
        counts = {}
        for issue in self._issues:
            v = getattr(issue, attr, u"")
            counts[v] = counts.get(v, 0) + 1
        return [(k, counts.get(k, 0)) for k in keys]

    def _count_top(self, attr, max_items=6):
        counts = {}
        for issue in self._issues:
            v = getattr(issue, attr, u"") or u"(sin asignar)"
            counts[v] = counts.get(v, 0) + 1
        return sorted(counts.items(), key=lambda x: x[1], reverse=True)[:max_items]

    # ---- grafico dona --------------------------------------------------------
    def _draw_donut(self):
        import math
        try:
            from System.Windows import Point as _Pt, Size as _Sz
            from System.Windows.Media import (
                PathGeometry    as _PG, PathFigure  as _PF,
                ArcSegment      as _AS, LineSegment as _LS,
                SolidColorBrush as _SB, ColorConverter as _CC,
                SweepDirection  as _SD, FillRule    as _FR,
            )

            counts = {u"Abierto": 0, u"En revision": 0, u"Resuelto": 0}
            _norm = {u"En revisi\xf3n": u"En revision"}  # normalizar "En revisión" -> "En revision"
            for issue in self._issues:
                st = _norm.get(issue.Estado, issue.Estado)
                if st in counts:
                    counts[st] += 1

            self._lbl_ab.Text  = str(counts[u"Abierto"])
            self._lbl_rev.Text = str(counts[u"En revision"])
            self._lbl_res.Text = str(counts[u"Resuelto"])

            total = sum(counts.values())
            cx, cy      = 65.0, 65.0
            R_out, R_in = 60.0, 34.0

            _segs = [
                (u"Abierto",     self._path_ab,  u"#1A3D70"),
                (u"En revision", self._path_rev, u"#3D1A6B"),
                (u"Resuelto",    self._path_res, u"#1A4F2A"),
            ]

            def _pt(r, deg):
                rad = math.radians(deg)
                return _Pt(cx + r * math.cos(rad), cy + r * math.sin(rad))

            def _full_ring(color):
                def _cfig(r):
                    fig = _PF()
                    fig.StartPoint = _Pt(cx, cy - r)
                    fig.IsClosed = True
                    fig.Segments.Add(_AS(_Pt(cx, cy + r),
                                         _Sz(r, r), 0.0, False, _SD.Clockwise, True))
                    fig.Segments.Add(_AS(_Pt(cx, cy - r),
                                         _Sz(r, r), 0.0, False, _SD.Clockwise, True))
                    return fig
                g = _PG()
                g.FillRule = _FR.EvenOdd
                g.Figures.Add(_cfig(R_out))
                g.Figures.Add(_cfig(R_in))
                return g, _SB(_CC.ConvertFromString(color))

            def _sector(start, sweep, color):
                if sweep <= 0:
                    return None
                end   = start + sweep
                large = sweep > 180.0
                p1 = _pt(R_out, start); p2 = _pt(R_out, end)
                p3 = _pt(R_in,  end);   p4 = _pt(R_in,  start)
                fig = _PF()
                fig.StartPoint = p1
                fig.IsClosed   = True
                fig.Segments.Add(_AS(p2, _Sz(R_out, R_out), 0.0,
                                     large, _SD.Clockwise, True))
                fig.Segments.Add(_LS(p3, True))
                fig.Segments.Add(_AS(p4, _Sz(R_in, R_in),  0.0,
                                     large, _SD.Counterclockwise, True))
                g = _PG()
                g.Figures.Add(fig)
                return g, _SB(_CC.ConvertFromString(color))

            if total == 0:
                for _, pe, _ in _segs:
                    pe.Data = None
                    pe.Fill = None
                return

            non_zero = [(n, c, p, col)
                        for (n, p, col), c in
                        zip(_segs, [counts[s[0]] for s in _segs])
                        if c > 0]

            if len(non_zero) == 1:
                n100, _, p100, col100 = non_zero[0]
                for n, p, col in _segs:
                    if n == n100:
                        p100.Data, p100.Fill = _full_ring(col100)
                    else:
                        p.Data = None
                        p.Fill = None
                return

            angle = -90.0
            for name, pe, color in _segs:
                cnt = counts.get(name, 0)
                if cnt == 0:
                    pe.Data = None
                    pe.Fill = None
                    continue
                sweep  = 360.0 * cnt / total
                result = _sector(angle, sweep, color)
                if result:
                    pe.Data, pe.Fill = result
                angle += sweep

        except Exception:
            pass

    # ---- grafico de barras horizontales --------------------------------------
    def _draw_bars(self, canvas, data, color_fn,
                   bar_max_w=175, label_w=105, bar_h=22, gap=10, max_bars=6):
        try:
            import System.Windows as _sw
            from System.Windows.Controls import Canvas as _Cv, TextBlock as _TB
            from System.Windows.Shapes   import Rectangle as _Rect
            from System.Windows.Media    import (
                SolidColorBrush as _SB, ColorConverter as _CC)

            canvas.Children.Clear()
            if not data:
                return

            items   = data[:max_bars]
            max_val = max(c for _, c in items) if items else 0

            for i, (label, count) in enumerate(items):
                y = float(i * (bar_h + gap))

                # Etiqueta
                tb = _TB()
                _d = label[:16] + (u"..." if len(label) > 16 else u"")
                tb.Text       = _d
                tb.Width      = float(label_w)
                tb.FontSize   = 11.0
                tb.Foreground = _SB(_CC.ConvertFromString(u"#C8E4EF"))
                _Cv.SetLeft(tb, 0.0)
                _Cv.SetTop(tb,  y + 3.0)
                canvas.Children.Add(tb)

                # Barra
                bw = 0.0
                if count > 0 and max_val > 0:
                    bw = max(4.0, float(bar_max_w) * count / max_val)
                    rect = _Rect()
                    rect.Width   = bw
                    rect.Height  = float(bar_h)
                    rect.RadiusX = 4.0
                    rect.RadiusY = 4.0
                    rect.Fill    = _SB(_CC.ConvertFromString(color_fn(label)))
                    _Cv.SetLeft(rect, float(label_w) + 8.0)
                    _Cv.SetTop(rect,  y)
                    canvas.Children.Add(rect)

                # Contador
                tb2 = _TB()
                tb2.Text       = str(count)
                tb2.FontSize   = 11.0
                tb2.FontWeight = _sw.FontWeights.Bold
                tb2.Foreground = _SB(_CC.ConvertFromString(u"#5BB8D4"))
                _Cv.SetLeft(tb2, float(label_w) + 8.0 + bw + 6.0)
                _Cv.SetTop(tb2,  y + 3.0)
                canvas.Children.Add(tb2)

        except Exception:
            pass

    # ---- cerrar con ESC ------------------------------------------------------
    def _on_key_down(self, sender, args):
        try:
            if args.Key == Key.Escape:
                self._win.Close()
        except Exception:
            pass

    def ShowDialog(self):
        self._win.ShowDialog()


# -- Ventana principal ---------------------------------------------------------
class IssuesViewerWindow(object):

    def __init__(self):
        self._issues_dir    = ISSUES_DIR
        self._IssueItem     = IssueItem
        self._DetAdjItem    = DetAdjuntoItem
        self._load_fn       = _load_all_issues
        self._OC         = ObservableCollection
        self._all_issues = []
        self._em         = EM_DASH
        self._n_ord      = N_ORDINAL
        self._logo_paths = _LOGO_PATHS

        # Revit context
        self._doc        = _doc
        self._uidoc      = _uidoc
        self._selected_issue  = None
        self._updating_estado = False
        self._updating_prioridad = False
        self._updating_asignado = False
        self._DashboardWindow = DashboardWindow
        # Revit API type refs (para handlers WPF)
        self._FEC   = FilteredElementCollector
        self._V3D   = View3D
        self._VFT   = ViewFamilyType
        self._VF    = ViewFamily
        self._XYZ   = XYZ
        self._VO3D  = ViewOrientation3D
        self._Tx    = Transaction

        win = XamlReader.Parse(XAML)
        self._win = win

        self._grid             = win.FindName("GridIssues")
        self._cmb_proyecto     = win.FindName("CmbFiltroProyecto")
        self._cmb_estado       = win.FindName("CmbFiltroEstado")
        self._cmb_prioridad    = win.FindName("CmbFiltroPrioridad")
        self._txt_total        = win.FindName("TxtTotal")
        self._txt_servidor     = win.FindName("TxtServidor")
        self._txt_status       = win.FindName("TxtStatus")

        self._det_placeholder  = win.FindName("DetPlaceholder")
        self._det_contenido    = win.FindName("DetContenido")
        self._img_detail       = win.FindName("ImgDetail")
        self._det_img_ph       = win.FindName("DetImgPlaceholder")
        self._det_badge_num    = win.FindName("DetBadgeNum")
        self._det_badge_proy   = win.FindName("DetBadgeProyecto")
        self._det_titulo       = win.FindName("DetTitulo")
        self._det_bd_prioridad = win.FindName("DetBdPrioridad")
        self._det_prioridad    = win.FindName("DetPrioridad")
        self._det_bd_estado    = win.FindName("DetBdEstado")
        self._det_estado       = win.FindName("DetEstado")
        self._det_descripcion  = win.FindName("DetDescripcion")
        self._det_disciplinas  = win.FindName("DetDisciplinas")
        self._det_reportado    = win.FindName("DetReportadoPor")
        self._det_asignado     = win.FindName("DetAsignadoA")
        self._det_fecha        = win.FindName("DetFecha")
        self._det_autor        = win.FindName("DetAutor")
        self._det_elementos      = win.FindName("DetElementos")
        self._det_vista          = win.FindName("DetVista")
        self._btn_viewpoint      = win.FindName("BtnIrViewpoint")
        self._img_logo           = win.FindName("ImgLogo")
        self._grid_placeholder   = win.FindName("GridPlaceholder")
        self._lst_det_adjuntos   = win.FindName("LstDetAdjuntos")
        self._det_sin_adjuntos   = win.FindName("DetSinAdjuntos")
        self._btn_abrir_adjunto  = win.FindName("BtnAbrirAdjunto")
        self._cmb_det_estado     = win.FindName("CmbDetEstado")
        self._txt_estado_confirm = win.FindName("TxtEstadoConfirm")
        self._cmb_det_prioridad  = win.FindName("CmbDetPrioridad")
        self._txt_prioridad_confirm = win.FindName("TxtPrioridadConfirm")
        self._cmb_det_asignado   = win.FindName("CmbDetAsignado")
        self._txt_asignado_confirm = win.FindName("TxtAsignadoConfirm")

        self._txt_servidor.Text = self._issues_dir

        self._load_logo()
        self._all_issues = _load_all_issues(self._issues_dir)
        self._setup_proyecto_filter()
        self._apply_filters()

        win.FindName("BtnRefresh").Click        += RoutedEventHandler(self._on_refresh)
        win.FindName("BtnLimpiarFiltros").Click += RoutedEventHandler(self._on_limpiar_filtros)
        win.FindName("BtnDashboard").Click      += RoutedEventHandler(self._on_dashboard)
        win.FindName("BtnManual").Click         += RoutedEventHandler(self._on_manual)
        self._cmb_proyecto.SelectionChanged     += SelectionChangedEventHandler(self._on_filter_changed)
        self._cmb_estado.SelectionChanged       += SelectionChangedEventHandler(self._on_filter_changed)
        self._cmb_prioridad.SelectionChanged    += SelectionChangedEventHandler(self._on_filter_changed)
        self._grid.SelectionChanged             += SelectionChangedEventHandler(self._on_issue_selected)
        self._btn_viewpoint.Click               += RoutedEventHandler(self._on_ir_viewpoint)
        self._btn_abrir_adjunto.Click           += RoutedEventHandler(self._on_abrir_adjunto)
        self._lst_det_adjuntos.SelectionChanged += SelectionChangedEventHandler(self._on_adjunto_det_selected)
        self._cmb_det_estado.SelectionChanged   += SelectionChangedEventHandler(self._on_estado_changed)
        self._cmb_det_prioridad.SelectionChanged += SelectionChangedEventHandler(self._on_prioridad_changed)
        self._cmb_det_asignado.SelectionChanged += SelectionChangedEventHandler(self._on_asignado_changed)
        self._win.KeyDown                       += KeyEventHandler(self._on_key_down)

    # -- Logo ------------------------------------------------------------------
    def _load_logo(self):
        import os
        import System
        try:
            if not self._img_logo:
                return
            for logo_path in self._logo_paths:
                if os.path.exists(logo_path):
                    bmp = BitmapImage()
                    bmp.BeginInit()
                    bmp.UriSource   = Uri(logo_path, UriKind.Absolute)
                    bmp.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad
                    bmp.EndInit()
                    bmp.Freeze()
                    self._img_logo.Source = bmp
                    break
        except Exception:
            pass

    # -- Setup filtros ---------------------------------------------------------
    def _setup_proyecto_filter(self):
        try:
            from System.Windows.Controls import ComboBoxItem
            self._cmb_proyecto.Items.Clear()
            item_todos = ComboBoxItem()
            item_todos.Content = u"Todos"
            self._cmb_proyecto.Items.Add(item_todos)
            proyectos = sorted(set(i.Proyecto for i in self._all_issues))
            for p in proyectos:
                item = ComboBoxItem()
                item.Content = p
                self._cmb_proyecto.Items.Add(item)
            # Por defecto: filtrar por proyecto actual si existe
            current_project = _get_project_folder_name(self._doc) if self._doc else None
            if current_project and proyectos:
                for idx, p in enumerate(proyectos):
                    if p == current_project:
                        self._cmb_proyecto.SelectedIndex = idx + 1  # +1 porque 0 es "Todos"
                        break
                else:
                    self._cmb_proyecto.SelectedIndex = 0
            else:
                self._cmb_proyecto.SelectedIndex = 0
        except Exception:
            pass

    # -- Aplicar filtros -------------------------------------------------------
    def _apply_filters(self):
        try:
            import System.Windows as _sw

            proyecto  = self._get_combo_text(self._cmb_proyecto)
            estado    = self._get_combo_text(self._cmb_estado)
            prioridad = self._get_combo_text(self._cmb_prioridad)

            col = self._OC[object]()
            count = 0
            for issue in self._all_issues:
                if proyecto  and proyecto  != u"Todos" and issue.Proyecto  != proyecto:
                    continue
                if estado    and estado    != u"Todos" and issue.Estado    != estado:
                    continue
                if prioridad and prioridad != u"Todos" and issue.Prioridad != prioridad:
                    continue
                col.Add(issue)
                count += 1

            self._grid.ItemsSource = col
            _suf = u"s" if count != 1 else u""
            self._txt_total.Text = u"{} incidencia{}".format(count, _suf)
            self._clear_detail()

            # Mostrar placeholder si no hay resultados
            _total = len(self._all_issues)
            if _total == 0:
                self._grid_placeholder.Visibility = _sw.Visibility.Visible
                self._txt_status.Text = (
                    u"Aun no se han creado incidencias. "
                    u"Usa BIM Issue para crear la primera."
                )
            elif count == 0:
                self._grid_placeholder.Visibility = _sw.Visibility.Visible
                self._txt_status.Text = (
                    u"Ningun resultado con los filtros aplicados "
                    u"({} incidencia{} en total).".format(_total, u"s" if _total != 1 else u"")
                )
            else:
                self._grid_placeholder.Visibility = _sw.Visibility.Collapsed
                self._txt_status.Text = u"Mostrando {} de {} incidencias.".format(
                    count, _total
                )
        except Exception as ex:
            self._txt_status.Text = u"Error al filtrar: " + str(ex)

    # -- Colores del combo de estado -------------------------------------------
    _ESTADO_COLORS = {
        u"Abierto":     (u"#1A3D70", u"#7EB8FF"),
        u"En revision": (u"#3D1A6B", u"#C97EFF"),
        u"Resuelto":    (u"#1A4F2A", u"#6BCC8E"),
    }

    # -- Colores del combo de prioridad ---------------------------------------
    _PRIORIDAD_COLORS = {
        u"Critica": (u"#6B1A1A", u"#FFB4B4"),
        u"Alta":    (u"#7A3B0A", u"#FFB76B"),
        u"Media":   (u"#5B4E08", u"#FFE47A"),
        u"Baja":    (u"#1A4F2A", u"#6BCC8E"),
    }

    def _update_estado_combo_color(self, estado):
        try:
            from System.Windows.Media import SolidColorBrush, ColorConverter
            bg, fg = self._ESTADO_COLORS.get(estado, (u"#1A3D52", u"#C8E4EF"))
            self._cmb_det_estado.Background = SolidColorBrush(
                ColorConverter.ConvertFromString(bg))
            self._cmb_det_estado.Foreground = SolidColorBrush(
                ColorConverter.ConvertFromString(fg))
        except Exception:
            pass

    def _update_prioridad_combo_color(self, prioridad):
        try:
            from System.Windows.Media import SolidColorBrush, ColorConverter
            bg, fg = self._PRIORIDAD_COLORS.get(prioridad, (u"#1A3D52", u"#C8E4EF"))
            self._cmb_det_prioridad.Background = SolidColorBrush(
                ColorConverter.ConvertFromString(bg))
            self._cmb_det_prioridad.Foreground = SolidColorBrush(
                ColorConverter.ConvertFromString(fg))
        except Exception:
            pass

    # -- Guardar cambio de prioridad -------------------------------------------
    def _on_prioridad_changed(self, sender, args):
        if self._updating_prioridad:
            return
        import os
        import json
        import System
        import System.IO as _sio
        import System.Windows as _sw
        try:
            issue = self._selected_issue
            if issue is None:
                return
            new_prioridad = self._get_combo_text(self._cmb_det_prioridad)
            if not new_prioridad or new_prioridad == issue.Prioridad:
                return

            # Persistir en issue.json
            json_path = os.path.join(issue.IssueDir, "issue.json")
            raw  = _sio.File.ReadAllText(json_path, System.Text.Encoding.UTF8)
            data = json.loads(raw)
            data["prioridad"] = new_prioridad
            new_raw   = json.dumps(data, ensure_ascii=False, indent=2)
            raw_bytes = System.Text.Encoding.UTF8.GetBytes(new_raw)
            _sio.File.WriteAllBytes(json_path, raw_bytes)

            # Actualizar objeto en memoria
            issue.Prioridad      = new_prioridad
            issue._data["prioridad"] = new_prioridad

            # Actualizar badge del panel de detalle
            _pri_colors = {
                u"Critica": (u"#6B1A1A", u"#FFB4B4"),
                u"Alta":    (u"#7A3B0A", u"#FFB76B"),
                u"Media":   (u"#5B4E08", u"#FFE47A"),
                u"Baja":    (u"#1A4F2A", u"#6BCC8E"),
            }
            self._set_badge(
                self._det_bd_prioridad, self._det_prioridad,
                new_prioridad, _pri_colors
            )

            # Actualizar color del combo
            self._update_prioridad_combo_color(new_prioridad)

            # Refrescar fila del DataGrid
            try:
                self._grid.Items.Refresh()
            except Exception:
                pass

            # Confirmacion visual
            self._txt_prioridad_confirm.Visibility = _sw.Visibility.Visible
            self._txt_status.Text = (
                u"Prioridad actualizada a '" + new_prioridad + u"' y guardada."
            )
        except Exception as ex:
            self._txt_status.Text = u"Error al guardar prioridad: " + str(ex)

    # -- Guardar cambio de estado ----------------------------------------------
    def _on_estado_changed(self, sender, args):
        if self._updating_estado:
            return
        import os
        import json
        import System
        import System.IO as _sio
        import System.Windows as _sw
        try:
            issue = self._selected_issue
            if issue is None:
                return
            new_estado = self._get_combo_text(self._cmb_det_estado)
            if not new_estado or new_estado == issue.Estado:
                return

            # Persistir en issue.json
            json_path = os.path.join(issue.IssueDir, "issue.json")
            raw  = _sio.File.ReadAllText(json_path, System.Text.Encoding.UTF8)
            data = json.loads(raw)
            data["estado"] = new_estado
            new_raw   = json.dumps(data, ensure_ascii=False, indent=2)
            raw_bytes = System.Text.Encoding.UTF8.GetBytes(new_raw)
            _sio.File.WriteAllBytes(json_path, raw_bytes)

            # Actualizar objeto en memoria
            issue.Estado      = new_estado
            issue._data["estado"] = new_estado

            # Actualizar badge del panel de detalle
            self._set_badge(
                self._det_bd_estado, self._det_estado,
                new_estado, self._ESTADO_COLORS
            )

            # Actualizar color del combo
            self._update_estado_combo_color(new_estado)

            # Refrescar fila del DataGrid
            try:
                self._grid.Items.Refresh()
            except Exception:
                pass

            # Confirmacion visual
            self._txt_estado_confirm.Visibility = _sw.Visibility.Visible
            self._txt_status.Text = (
                u"Estado actualizado a '" + new_estado + u"' y guardado."
            )
        except Exception as ex:
            self._txt_status.Text = u"Error al guardar estado: " + str(ex)

    # -- Guardar cambio de asignado a -------------------------------------------
    def _on_asignado_changed(self, sender, args):
        if self._updating_asignado:
            return
        import os
        import json
        import System
        import System.IO as _sio
        import System.Windows as _sw
        try:
            issue = self._selected_issue
            if issue is None:
                return
            sel = self._cmb_det_asignado.SelectedItem
            if sel is None:
                return
            if not isinstance(sel, PersonaItem):
                return
            new_nombre = (sel.Nombre or u"").strip()
            new_email = (sel.Email or u"").strip()
            if new_nombre == issue.AsignadoA:
                return

            # Persistir en issue.json
            json_path = os.path.join(issue.IssueDir, "issue.json")
            raw  = _sio.File.ReadAllText(json_path, System.Text.Encoding.UTF8)
            data = json.loads(raw)
            data["asignado_a"] = {"nombre": new_nombre, "email": new_email}
            new_raw   = json.dumps(data, ensure_ascii=False, indent=2)
            raw_bytes = System.Text.Encoding.UTF8.GetBytes(new_raw)
            _sio.File.WriteAllBytes(json_path, raw_bytes)

            # Actualizar objeto en memoria
            issue.AsignadoA = new_nombre
            issue._data["asignado_a"] = {"nombre": new_nombre, "email": new_email}

            # Actualizar TextBlock del panel de detalle
            self._det_asignado.Text = new_nombre

            # Refrescar fila del DataGrid
            try:
                self._grid.Items.Refresh()
            except Exception:
                pass

            # Confirmacion visual
            self._txt_asignado_confirm.Visibility = _sw.Visibility.Visible
            self._txt_status.Text = (
                u"Asignado actualizado a '" + new_nombre + u"' y guardado."
            )
        except Exception as ex:
            self._txt_status.Text = u"Error al guardar asignado: " + str(ex)

    # -- Limpiar detalle -------------------------------------------------------
    def _clear_detail(self):
        try:
            import System.Windows as _sw
            self._det_placeholder.Visibility = _sw.Visibility.Visible
            self._det_contenido.Visibility   = _sw.Visibility.Collapsed
            self._img_detail.Visibility      = _sw.Visibility.Collapsed
            self._det_img_ph.Visibility      = _sw.Visibility.Visible
            self._btn_viewpoint.IsEnabled    = False
            self._btn_abrir_adjunto.IsEnabled = False
            self._lst_det_adjuntos.ItemsSource = None
            self._lst_det_adjuntos.Visibility  = _sw.Visibility.Collapsed
            self._det_sin_adjuntos.Visibility  = _sw.Visibility.Visible
            self._cmb_det_estado.IsEnabled     = False
            self._cmb_det_prioridad.IsEnabled  = False
            self._cmb_det_asignado.IsEnabled   = False
            self._updating_estado = True
            self._updating_prioridad = True
            self._updating_asignado = True
            self._cmb_det_estado.SelectedIndex = -1
            self._cmb_det_prioridad.SelectedIndex = -1
            self._cmb_det_asignado.SelectedIndex = -1
            self._updating_estado = False
            self._updating_prioridad = False
            self._updating_asignado = False
            self._txt_estado_confirm.Visibility = _sw.Visibility.Collapsed
            self._txt_prioridad_confirm.Visibility = _sw.Visibility.Collapsed
            self._txt_asignado_confirm.Visibility = _sw.Visibility.Collapsed
            self._selected_issue             = None
        except Exception:
            pass

    # -- Mostrar detalle -------------------------------------------------------
    def _show_detail(self, issue):
        try:
            import System.Windows as _sw
            import System.Windows.Media as _swm

            self._det_placeholder.Visibility = _sw.Visibility.Collapsed
            self._det_contenido.Visibility   = _sw.Visibility.Visible

            # Screenshot
            ss = issue.ScreenshotPath
            if ss:
                try:
                    import System.IO as _sio
                    if _sio.File.Exists(ss):
                        bmp = BitmapImage()
                        bmp.BeginInit()
                        bmp.UriSource   = Uri(ss, UriKind.Absolute)
                        bmp.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad
                        bmp.EndInit()
                        bmp.Freeze()
                        self._img_detail.Source     = bmp
                        self._img_detail.Visibility = _sw.Visibility.Visible
                        self._det_img_ph.Visibility = _sw.Visibility.Collapsed
                    else:
                        self._img_detail.Visibility = _sw.Visibility.Collapsed
                        self._det_img_ph.Visibility = _sw.Visibility.Visible
                except Exception:
                    self._img_detail.Visibility = _sw.Visibility.Collapsed
                    self._det_img_ph.Visibility = _sw.Visibility.Visible
            else:
                self._img_detail.Visibility = _sw.Visibility.Collapsed
                self._det_img_ph.Visibility = _sw.Visibility.Visible

            # Textos
            self._det_badge_num.Text  = u"Incidencia N" + self._n_ord + u" {}".format(issue.Numero)
            self._det_badge_proy.Text = u"| {}".format(issue.Proyecto)
            self._det_titulo.Text     = issue.Titulo
            self._det_descripcion.Text = issue._data.get("descripcion", self._em)

            discs = issue._data.get("disciplinas", [])
            self._det_disciplinas.Text = (
                u", ".join(discs) if discs else self._em
            )

            self._det_reportado.Text = issue.ReportadoPor
            self._det_asignado.Text  = issue.AsignadoA
            self._det_fecha.Text     = issue.Fecha
            self._det_autor.Text     = issue._data.get("autor", self._em)

            elems = issue._data.get("elementos", [])
            self._det_elementos.Text = (
                u"{} elemento(s) vinculado(s)".format(len(elems)) if elems
                else u"Sin elementos vinculados"
            )

            vp          = issue._data.get("viewpoint", {})
            cam         = vp.get("camera")  if vp else None
            view_id_int = vp.get("view_id") if vp else None
            if vp:
                _vname = vp.get("view_name", self._em)
                _vtype = vp.get("view_type", self._em)
                _vid   = u"  [ID {}]".format(view_id_int) if view_id_int else u""
                self._det_vista.Text = u"{} ({}){}" .format(_vname, _vtype, _vid)
            else:
                self._det_vista.Text = self._em

            # Adjuntos
            adjuntos = issue._data.get("adjuntos", [])
            if adjuntos:
                import os as _os
                _adj_col = self._OC[object]()
                for _fname in adjuntos:
                    _ruta = _os.path.join(issue.IssueDir, "adjuntos", _fname)
                    _adj_col.Add(self._DetAdjItem(_fname, _ruta))
                self._lst_det_adjuntos.ItemsSource = _adj_col
                self._lst_det_adjuntos.Visibility  = _sw.Visibility.Visible
                self._det_sin_adjuntos.Visibility  = _sw.Visibility.Collapsed
            else:
                self._lst_det_adjuntos.ItemsSource = None
                self._lst_det_adjuntos.Visibility  = _sw.Visibility.Collapsed
                self._det_sin_adjuntos.Visibility  = _sw.Visibility.Visible
            self._btn_abrir_adjunto.IsEnabled = False

            # Boton activo si hay camara 3D o view_id (cualquier tipo de vista)
            _has_nav = (cam is not None or view_id_int is not None) and self._doc is not None
            self._btn_viewpoint.IsEnabled = _has_nav

            # Combo de cambio de prioridad
            self._updating_prioridad = True
            for _i in range(self._cmb_det_prioridad.Items.Count):
                _it = self._cmb_det_prioridad.Items[_i]
                if hasattr(_it, 'Content') and str(_it.Content) == issue.Prioridad:
                    self._cmb_det_prioridad.SelectedIndex = _i
                    break
            self._cmb_det_prioridad.IsEnabled = True
            self._txt_prioridad_confirm.Visibility = _sw.Visibility.Collapsed
            self._updating_prioridad = False
            self._update_prioridad_combo_color(issue.Prioridad)

            # Combo de cambio de estado
            self._updating_estado = True
            for _i in range(self._cmb_det_estado.Items.Count):
                _it = self._cmb_det_estado.Items[_i]
                if hasattr(_it, 'Content') and str(_it.Content) == issue.Estado:
                    self._cmb_det_estado.SelectedIndex = _i
                    break
            self._cmb_det_estado.IsEnabled = True
            self._txt_estado_confirm.Visibility = _sw.Visibility.Collapsed
            self._updating_estado = False
            self._update_estado_combo_color(issue.Estado)

            # Combo de cambio de asignado a
            _personas = [PersonaItem(EM_DASH, u"")] + sorted(
                _load_personas(PERSONAS_FILE), key=lambda p: (p.Nombre or u"").lower()
            )
            self._cmb_det_asignado.ItemsSource = _personas
            self._updating_asignado = True
            _asi_nom = (issue.AsignadoA or u"").strip()
            for _i, _p in enumerate(_personas):
                if (_p.Nombre or u"").strip() == _asi_nom:
                    self._cmb_det_asignado.SelectedIndex = _i
                    break
            else:
                self._cmb_det_asignado.SelectedIndex = -1
            self._cmb_det_asignado.IsEnabled = True
            self._txt_asignado_confirm.Visibility = _sw.Visibility.Collapsed
            self._updating_asignado = False

            # Colores de badges
            _pri_colors = {
                u"Critica": (u"#6B1A1A", u"#FFB4B4"),
                u"Alta":    (u"#7A3B0A", u"#FFB76B"),
                u"Media":   (u"#5B4E08", u"#FFE47A"),
                u"Baja":    (u"#1A4F2A", u"#6BCC8E"),
            }
            _est_colors = {
                u"Abierto":     (u"#1A3D70", u"#7EB8FF"),
                u"En revision": (u"#3D1A6B", u"#C97EFF"),
                u"Resuelto":    (u"#1A4F2A", u"#6BCC8E"),
            }
            self._set_badge(
                self._det_bd_prioridad, self._det_prioridad,
                issue.Prioridad, _pri_colors
            )
            self._set_badge(
                self._det_bd_estado, self._det_estado,
                issue.Estado, _est_colors
            )
        except Exception as ex:
            self._txt_status.Text = u"Error al mostrar detalle: " + str(ex)

    def _set_badge(self, border, label, value, color_map):
        try:
            from System.Windows.Media import SolidColorBrush, ColorConverter
            bg, fg = color_map.get(value, (u"#1A3D52", u"#C8E4EF"))
            border.Background = SolidColorBrush(ColorConverter.ConvertFromString(bg))
            label.Foreground  = SolidColorBrush(ColorConverter.ConvertFromString(fg))
            label.Text = value
        except Exception:
            label.Text = value

    # -- Helpers ---------------------------------------------------------------
    @staticmethod
    def _get_combo_text(combo):
        try:
            item = combo.SelectedItem
            if item is None:
                return u""
            content = getattr(item, "Content", None)
            return str(content) if content is not None else str(item)
        except Exception:
            return u""

    # -- Handlers --------------------------------------------------------------
    def _on_dashboard(self, sender, args):
        try:
            # Siempre usar TODAS las incidencias para el dashboard (no las filtradas)
            items = list(self._all_issues)
            if not items:
                self._txt_status.Text = u"No hay incidencias para mostrar en el dashboard."
                return
            dash = self._DashboardWindow(items)
            dash.ShowDialog()
        except Exception as ex:
            self._txt_status.Text = u"Error al abrir dashboard: " + str(ex)

    def _on_manual(self, sender, args):
        """Abre el manual de usuario en el navegador predeterminado."""
        try:
            _script_dir = os.path.dirname(os.path.abspath(__file__))
            manual_path = os.path.join(_script_dir, "manual_usuario.html")
            if os.path.exists(manual_path):
                abs_path = os.path.abspath(manual_path)
                os.startfile(abs_path)
                self._txt_status.Text = u"Manual abierto."
            else:
                MessageBox.Show(
                    u"No se encontró el archivo manual_usuario.html.\n\nRuta esperada:\n{}".format(manual_path),
                    u"Manual no encontrado",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                )
        except Exception as ex:
            self._txt_status.Text = u"Error al abrir manual: " + str(ex)

    def _on_refresh(self, sender, args):
        try:
            self._all_issues = self._load_fn(self._issues_dir)
            self._setup_proyecto_filter()
            self._apply_filters()
            self._txt_status.Text = u"Actualizado. {} incidencias cargadas.".format(
                len(self._all_issues)
            )
        except Exception as ex:
            self._txt_status.Text = u"Error al actualizar: " + str(ex)

    def _on_limpiar_filtros(self, sender, args):
        try:
            self._cmb_proyecto.SelectedIndex  = 0
            self._cmb_estado.SelectedIndex    = 0
            self._cmb_prioridad.SelectedIndex = 0
        except Exception:
            pass

    def _on_filter_changed(self, sender, args):
        self._apply_filters()

    def _on_issue_selected(self, sender, args):
        try:
            issue = self._grid.SelectedItem
            if issue and isinstance(issue, self._IssueItem):
                self._selected_issue = issue
                self._show_detail(issue)
                self._txt_status.Text = (
                    u"Incidencia N" + self._n_ord +
                    u" {} - {}".format(issue.Numero, issue.Titulo)
                )
            else:
                self._selected_issue = None
                self._clear_detail()
        except Exception as ex:
            self._txt_status.Text = u"Error: " + str(ex)

    def _on_adjunto_det_selected(self, sender, args):
        try:
            self._btn_abrir_adjunto.IsEnabled = (
                self._lst_det_adjuntos.SelectedItem is not None
            )
        except Exception:
            pass

    def _on_abrir_adjunto(self, sender, args):
        try:
            import os as _os
            sel = self._lst_det_adjuntos.SelectedItem
            if sel is None:
                return
            ruta = sel.Ruta
            if _os.path.exists(ruta):
                _os.startfile(ruta)
                self._txt_status.Text = u"Abriendo: " + sel.Nombre
            else:
                self._txt_status.Text = (
                    u"Archivo no encontrado: " + ruta
                )
        except Exception as ex:
            self._txt_status.Text = u"Error al abrir adjunto: " + str(ex)

    def _on_ir_viewpoint(self, sender, args):
        try:
            from Autodesk.Revit.DB import (
                FilteredElementCollector as _FEC,
                View3D                  as _V3D,
                View                    as _View,
                ViewFamilyType          as _VFT,
                ViewFamily              as _VF,
                XYZ                     as _XYZ,
                ViewOrientation3D       as _VO3D,
                Transaction             as _Tx,
                ElementId               as _EId,
            )

            issue = self._selected_issue
            if issue is None:
                return

            doc   = self._doc
            uidoc = self._uidoc
            if doc is None or uidoc is None:
                self._txt_status.Text = u"No hay documento activo en Revit."
                return

            vp          = issue._data.get("viewpoint", {})
            cam         = vp.get("camera")  if vp else None
            view_id_int = vp.get("view_id") if vp else None

            # ----------------------------------------------------------------
            # CASO 1: Vista 3D con datos de camara
            #   -> regenerar "Gestion de incidencias" con la orientacion guardada
            # ----------------------------------------------------------------
            if cam is not None:
                VIEW_NAME = u"Gestion de incidencias"

                view_gestion = None
                for v in _FEC(doc).OfClass(_V3D):
                    if not v.IsTemplate and v.Name == VIEW_NAME:
                        view_gestion = v
                        break

                with _Tx(doc, u"Ir a incidencia") as t:
                    t.Start()
                    try:
                        if view_gestion is None:
                            vft = None
                            for vft_item in _FEC(doc).OfClass(_VFT):
                                if vft_item.ViewFamily == _VF.ThreeDimensional:
                                    vft = vft_item
                                    break
                            if vft is None:
                                t.RollBack()
                                self._txt_status.Text = u"No se encontro un tipo de vista 3D."
                                return
                            view_gestion = _V3D.CreateIsometric(doc, vft.Id)
                            view_gestion.Name = VIEW_NAME

                        eye = _XYZ(cam["eye"][0],     cam["eye"][1],     cam["eye"][2])
                        fwd = _XYZ(cam["forward"][0],  cam["forward"][1], cam["forward"][2])
                        up  = _XYZ(cam["up"][0],       cam["up"][1],      cam["up"][2])
                        ori = _VO3D(eye, up, fwd)
                        view_gestion.SetOrientation(ori)

                        # Restaurar crop box (zoom exacto).
                        # Se lee DESPUES de SetOrientation para tener el Transform correcto.
                        cb_data = cam.get("crop_box")
                        if cb_data:
                            try:
                                _cb = view_gestion.CropBox
                                _cb.Min = _XYZ(
                                    cb_data["min"][0],
                                    cb_data["min"][1],
                                    cb_data["min"][2]
                                )
                                _cb.Max = _XYZ(
                                    cb_data["max"][0],
                                    cb_data["max"][1],
                                    cb_data["max"][2]
                                )
                                view_gestion.CropBox = _cb
                                view_gestion.CropBoxActive = cb_data.get("active", False)
                            except Exception:
                                pass

                        t.Commit()
                    except Exception as ex_inner:
                        t.RollBack()
                        self._txt_status.Text = u"Error en transaccion: " + str(ex_inner)
                        return

                try:
                    uidoc.ActiveView = view_gestion
                except Exception:
                    pass
                self._win.Close()
                return

            # ----------------------------------------------------------------
            # CASO 2: Vista 2D (o 3D sin camara) con view_id guardado
            #   -> navegar directamente a la vista por su ElementId
            # ----------------------------------------------------------------
            if view_id_int is not None:
                target_view = None
                try:
                    elem = doc.GetElement(_EId(int(view_id_int)))
                    if elem is not None and isinstance(elem, _View):
                        target_view = elem
                except Exception:
                    pass

                if target_view is None:
                    self._txt_status.Text = (
                        u"La vista original (ID {}) ya no existe en el documento.".format(
                            view_id_int)
                    )
                    return

                try:
                    uidoc.ActiveView = target_view
                except Exception:
                    pass
                self._win.Close()
                return

            # ----------------------------------------------------------------
            # CASO 3: Sin datos de navegacion
            # ----------------------------------------------------------------
            self._txt_status.Text = u"La incidencia no tiene datos de vista guardados."

        except Exception as ex:
            self._txt_status.Text = u"Error al ir al punto de vista: " + str(ex)

    def _on_key_down(self, sender, args):
        try:
            if args.Key == Key.Escape:
                self._win.Close()
        except Exception:
            pass

    def ShowDialog(self):
        self._win.ShowDialog()


# -- Entry point ---------------------------------------------------------------
try:
    w = IssuesViewerWindow()
    w.ShowDialog()
except Exception as ex:
    import traceback
    try:
        from pyrevit import forms as _pf
        _pf.alert(
            u"Error al abrir Ver Issues:\n\n{}\n\n{}".format(
                str(ex), traceback.format_exc()),
            title=u"Error")
    except Exception:
        from System.Windows import MessageBox as _MB
        _MB.Show(
            u"{}\n\n{}".format(str(ex), traceback.format_exc()),
            u"Error - VerIssues")
