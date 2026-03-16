# -*- coding: utf-8 -*-
"""
Personas — Gestiona el directorio de personas del equipo BIM.
Comparte el archivo personas.json con el botón BIM Issue.
"""

__title__ = "Personas"
__author__ = "pyRevit"
__doc__    = "Agrega, edita o elimina personas del directorio compartido con BIM Issue."

import os
import json
import io

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Collections")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Threading")

from System.Windows.Markup          import XamlReader
from System.Windows.Threading       import Dispatcher
from System.Windows                 import Window, MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows.Controls        import Button
from System.Windows.Input           import Key, KeyEventHandler, MouseButtonEventHandler
from System.Windows.Media           import VisualTreeHelper
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Media.Imaging   import BitmapImage
from System                         import Uri, UriKind, Action
import System.IO as sio
import System

ISSUES_DIR    = u"Y:\\00_SERVIDOR DE INCIDENCIAS"
PERSONAS_FILE = os.path.join(ISSUES_DIR, "personas.json")


# ── Modelo ──────────────────────────────────────────────────────────────────
class PersonaItem(object):
    def __init__(self, nombre, email=""):
        self.Nombre = nombre
        self.Email  = email

    def __str__(self):
        return self.Nombre

    def ToString(self):
        return self.Nombre


# ── Persistencia ─────────────────────────────────────────────────────────────
def _load_personas(personas_file):
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


def _save_personas(collection, issues_dir, personas_file):
    import json
    sio.Directory.CreateDirectory(issues_dir)
    data = [{"nombre": p.Nombre, "email": p.Email} for p in collection]
    sio.File.WriteAllText(
        personas_file,
        json.dumps(data, ensure_ascii=False, indent=2),
        System.Text.Encoding.UTF8,
    )


# ── XAML ─────────────────────────────────────────────────────────────────────
XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Arainco - Personas"
    Height="680" Width="680"
    MinHeight="500" MinWidth="560"
    WindowStartupLocation="CenterScreen"
    Background="#0A1C26"
    FontFamily="Segoe UI"
    ResizeMode="CanResize">

  <Window.Resources>
    <Style x:Key="Label" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#4A8BA6"/>
      <Setter Property="FontSize"   Value="10"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin"     Value="0,0,0,4"/>
    </Style>
    <Style x:Key="Field" TargetType="TextBox">
      <Setter Property="Background"      Value="#0A1C26"/>
      <Setter Property="Foreground"      Value="#C8E4EF"/>
      <Setter Property="CaretBrush"      Value="#5BB8D4"/>
      <Setter Property="BorderBrush"     Value="#1A3D52"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="10,8"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="6" Padding="{TemplateBinding Padding}">
              <Grid>
                <TextBlock x:Name="Wm" Text="{TemplateBinding Tag}"
                           Foreground="#2A5C75" FontStyle="Italic"
                           IsHitTestVisible="False" Visibility="Collapsed"
                           VerticalAlignment="Center"/>
                <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="Center"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <DataTrigger Binding="{Binding Text, RelativeSource={RelativeSource Self}}" Value="">
                <Setter TargetName="Wm" Property="Visibility" Value="Visible"/>
              </DataTrigger>
              <Trigger Property="IsFocused" Value="True">
                <Setter TargetName="Wm"  Property="Visibility" Value="Collapsed"/>
                <Setter Property="BorderBrush" Value="#5BB8D4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="Padding"         Value="18,9"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                    CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Opacity" Value="0.85"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Bd" Property="Opacity" Value="0.7"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
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
        <StackPanel DockPanel.Dock="Left" Orientation="Horizontal" VerticalAlignment="Center">
          <Image x:Name="ImgLogo" Width="80" Height="80" Margin="0,0,16,0"
                 Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality"/>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="DIRECTORIO DE PERSONAS" FontSize="18" FontWeight="Bold"
                       Foreground="#C8E4EF"/>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
              <Border Background="#1A4F6A" CornerRadius="8" Padding="8,2">
                <TextBlock x:Name="TxtPersonasCount" Foreground="#5BB8D4"
                           FontSize="11" FontWeight="Bold" Text="0 personas"/>
              </Border>
            </StackPanel>
          </StackPanel>
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- Formulario agregar / editar -->
    <Border Grid.Row="1" Background="#0F2535" BorderBrush="#1A3D52"
            BorderThickness="0,0,0,1" Padding="20,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Column="0" Margin="0,0,10,0">
          <TextBlock Text="NOMBRE *" Style="{StaticResource Label}"/>
          <TextBox x:Name="TxtNombre" Style="{StaticResource Field}"
                   Tag="Nombre completo"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Margin="0,0,10,0">
          <TextBlock Text="EMAIL" Style="{StaticResource Label}"/>
          <TextBox x:Name="TxtEmail" Style="{StaticResource Field}"
                   Tag="correo@empresa.com"/>
        </StackPanel>

        <Button x:Name="BtnAgregar" Grid.Column="2"
                Content="+ Agregar" VerticalAlignment="Bottom"
                Background="#5BB8D4" Foreground="#0A1C26"
                Style="{StaticResource Btn}" Margin="0,0,8,0"/>
        <Button x:Name="BtnActualizar" Grid.Column="3"
                Content="&#10003; Guardar cambio" VerticalAlignment="Bottom"
                Background="#1A4F6A" Foreground="#C8E4EF"
                Style="{StaticResource Btn}" Visibility="Collapsed"/>
        <Button x:Name="BtnRecargar" Grid.Column="4"
                Content="Recargar" VerticalAlignment="Bottom"
                Background="#1A4F6A" Foreground="#C8E4EF"
                Style="{StaticResource Btn}" Margin="8,0,0,0"
                ToolTip="Recarga la lista desde el archivo (después de importar)"/>
      </Grid>
    </Border>

    <!-- Tabla -->
    <Grid Grid.Row="2" Margin="20,16,20,0">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <DockPanel Grid.Row="0" Margin="0,0,0,8">
        <TextBlock DockPanel.Dock="Left" Text="PERSONAS REGISTRADAS"
                   Foreground="#4A8BA6" FontSize="10" FontWeight="SemiBold"
                   VerticalAlignment="Center"/>
        <TextBlock x:Name="TxtConteo" DockPanel.Dock="Right"
                   Foreground="#2A5570" FontSize="10" VerticalAlignment="Center"
                   HorizontalAlignment="Right"/>
      </DockPanel>

      <DataGrid Grid.Row="1" x:Name="GridPersonas"
                AutoGenerateColumns="False" IsReadOnly="True"
                SelectionMode="Single" CanUserAddRows="False"
                CanUserReorderColumns="False" CanUserResizeRows="False"
                Background="#081520" Foreground="#C8E4EF"
                BorderBrush="#1A3D52" BorderThickness="1"
                GridLinesVisibility="Horizontal"
                HorizontalGridLinesBrush="#132C3D"
                RowBackground="#081520"
                AlternatingRowBackground="#0B1F2E"
                ColumnHeaderHeight="32" RowHeight="40" FontSize="13">
        <DataGrid.Columns>
          <DataGridTextColumn Header="Nombre" Binding="{Binding Nombre}" Width="*"
                              ElementStyle="{StaticResource {x:Type TextBlock}}"/>
          <DataGridTextColumn Header="Email"  Binding="{Binding Email}"  Width="*"/>
          <DataGridTemplateColumn Header="" Width="130" CanUserResize="False">
            <DataGridTemplateColumn.CellTemplate>
              <DataTemplate>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center"
                            HorizontalAlignment="Center">
                  <Button x:Name="BtnEditar" Content="Editar"
                          Tag="{Binding}"
                          Background="#1A3D52" Foreground="#5BB8D4"
                          BorderThickness="0" Cursor="Hand"
                          FontSize="11" Padding="10,4" Margin="0,0,6,0">
                    <Button.Template>
                      <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                          <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="Bd" Property="Background" Value="#235E7D"/>
                          </Trigger>
                        </ControlTemplate.Triggers>
                      </ControlTemplate>
                    </Button.Template>
                  </Button>
                  <Button x:Name="BtnEliminar" Content="&#10005;"
                          Tag="{Binding}"
                          Background="#2A1A1A" Foreground="#D45B5B"
                          BorderThickness="0" Cursor="Hand"
                          FontSize="12" FontWeight="Bold" Padding="8,4">
                    <Button.Template>
                      <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                          <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="Bd" Property="Background" Value="#3D1A1A"/>
                          </Trigger>
                        </ControlTemplate.Triggers>
                      </ControlTemplate>
                    </Button.Template>
                  </Button>
                </StackPanel>
              </DataTemplate>
            </DataGridTemplateColumn.CellTemplate>
          </DataGridTemplateColumn>
        </DataGrid.Columns>
        <DataGrid.ColumnHeaderStyle>
          <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background"      Value="#1A3D52"/>
            <Setter Property="Foreground"      Value="#5BB8D4"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="12,0"/>
            <Setter Property="BorderBrush"     Value="#2A5570"/>
            <Setter Property="BorderThickness" Value="0,0,1,0"/>
          </Style>
        </DataGrid.ColumnHeaderStyle>
        <DataGrid.CellStyle>
          <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding"         Value="12,0"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Style.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter Property="Background" Value="#1A3D52"/>
                <Setter Property="Foreground" Value="#C8E4EF"/>
              </Trigger>
            </Style.Triggers>
          </Style>
        </DataGrid.CellStyle>
        <DataGrid.RowStyle>
          <Style TargetType="DataGridRow">
            <Style.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#0F2535"/>
              </Trigger>
            </Style.Triggers>
          </Style>
        </DataGrid.RowStyle>
      </DataGrid>
    </Grid>

    <!-- Footer -->
    <Border Grid.Row="3" Background="#0F2535" BorderBrush="#1A3D52"
            BorderThickness="0,1,0,0" Padding="20,12">
      <TextBlock x:Name="TxtEstado"
                 VerticalAlignment="Center" FontSize="11" Foreground="#4A8BA6"
                 Text="Selecciona una fila para editar o eliminar."/>
    </Border>
  </Grid>
</Window>
"""


# ── Ventana principal ─────────────────────────────────────────────────────────
class PersonasWindow(Window):

    def __init__(self):
        # Almacenar constantes como instancia (los handlers WPF no acceden al scope global)
        self._issues_dir    = ISSUES_DIR
        self._personas_file = PERSONAS_FILE

        win = XamlReader.Parse(XAML)
        self.Content               = win.Content
        self.Title                 = win.Title
        self.Width                 = win.Width
        self.Height                = win.Height
        self.MinWidth              = win.MinWidth
        self.MinHeight             = win.MinHeight
        self.Background            = win.Background
        self.FontFamily            = win.FontFamily
        self.WindowStartupLocation = win.WindowStartupLocation
        self.ResizeMode            = win.ResizeMode

        self._txt_nombre    = win.FindName("TxtNombre")
        self._txt_email     = win.FindName("TxtEmail")
        self._btn_agregar   = win.FindName("BtnAgregar")
        self._btn_actualizar= win.FindName("BtnActualizar")
        self._grid          = win.FindName("GridPersonas")
        self._txt_estado         = win.FindName("TxtEstado")
        self._txt_conteo         = win.FindName("TxtConteo")
        self._txt_personas_count = win.FindName("TxtPersonasCount")

        self._personas = ObservableCollection[object]()
        for p in _load_personas(self._personas_file):
            self._personas.Add(p)

        self._grid.ItemsSource = self._personas
        self._editing_item     = None

        self._update_conteo()
        self._load_logo(win)

        self._btn_agregar.Click    += self._on_agregar
        self._btn_actualizar.Click += self._on_actualizar
        self._btn_recargar = win.FindName("BtnRecargar")
        if self._btn_recargar:
            self._btn_recargar.Click += self._on_recargar
        self._grid.PreviewMouseLeftButtonDown += MouseButtonEventHandler(self._on_grid_preview_mouse_down)
        self.KeyDown               += KeyEventHandler(self._on_key_down)
        self._personas.CollectionChanged += lambda s, e: self._update_conteo()

    # ── Clic en botones Editar/Eliminar (via PreviewMouseDown) ─────────────────
    def _on_grid_preview_mouse_down(self, sender, args):
        try:
            dep = args.OriginalSource
            while dep is not None:
                if isinstance(dep, Button):
                    persona = dep.Tag
                    if isinstance(persona, PersonaItem):
                        name = getattr(dep, "Name", None) or ""
                        content = str(getattr(dep, "Content", "") or "")
                        if name == "BtnEditar" or content == "Editar":
                            self._on_editar_click(dep, None)
                            args.Handled = True
                        elif name == "BtnEliminar" or content in ("\u2715", "\u00D7", "x", "X"):
                            self._on_eliminar_click(dep, None)
                            args.Handled = True
                    break
                try:
                    dep = VisualTreeHelper.GetParent(dep)
                except Exception:
                    break
        except Exception:
            pass

    # ── Logo ─────────────────────────────────────────────────────────────────
    def _load_logo(self, win):
        try:
            import os
            logo_dir  = os.path.dirname(__file__)
            # Buscar logo en la propia carpeta, luego en 01_BIMIssue como fallback
            logo_path = os.path.join(logo_dir, "logo.png")
            if not os.path.exists(logo_path):
                # 01_BIMIssue está en Incidencias.stack (hermano), no en Incidencias2.stack
                panel_dir = os.path.dirname(os.path.dirname(logo_dir))
                bim_dir   = os.path.join(panel_dir, "Incidencias.stack", "01_BIMIssue.pushbutton")
                logo_path = os.path.join(bim_dir, "logo.png")
            if not os.path.exists(logo_path):
                return
            img_ctrl = win.FindName("ImgLogo")
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

    # ── Conteo ───────────────────────────────────────────────────────────────
    def _update_conteo(self):
        n = len(self._personas)
        _txt = u"{} persona{}".format(n, "s" if n != 1 else "")
        self._txt_conteo.Text = _txt
        if self._txt_personas_count:
            self._txt_personas_count.Text = _txt

    def _on_recargar(self, sender, args):
        """Recarga la lista de personas desde el archivo."""
        self._personas.Clear()
        for p in _load_personas(self._personas_file):
            self._personas.Add(p)
        self._set_estado(u"Lista recargada.")

    # ── Agregar ──────────────────────────────────────────────────────────────
    def _on_agregar(self, sender, args):
        nombre = self._txt_nombre.Text.strip()
        if not nombre:
            self._set_estado(u"El nombre es obligatorio.")
            self._txt_nombre.Focus()
            return
        email = self._txt_email.Text.strip()

        # Verificar duplicado por nombre
        for p in self._personas:
            if p.Nombre.lower() == nombre.lower():
                self._set_estado(u'Ya existe una persona con el nombre "{}".'.format(nombre))
                return

        self._personas.Add(PersonaItem(nombre, email))
        _save_personas(self._personas, self._issues_dir, self._personas_file)
        self._txt_nombre.Text = ""
        self._txt_email.Text  = ""
        self._txt_nombre.Focus()
        self._set_estado(u'"{}" agregado correctamente.'.format(nombre))

    # ── Editar (cargar en formulario) ────────────────────────────────────────
    def _on_editar_click(self, sender, args):
        persona = sender.Tag
        if not isinstance(persona, PersonaItem):
            return
        self._editing_item        = persona
        self._txt_nombre.Text     = persona.Nombre
        self._txt_email.Text      = persona.Email
        self._btn_agregar.Visibility    = System.Windows.Visibility.Collapsed
        self._btn_actualizar.Visibility = System.Windows.Visibility.Visible
        self._txt_nombre.Focus()
        self._set_estado(u'Editando: "{}". Modifica los campos y pulsa Guardar cambio.'.format(persona.Nombre))

    # ── Guardar cambio ────────────────────────────────────────────────────────
    def _on_actualizar(self, sender, args):
        nombre = self._txt_nombre.Text.strip()
        if not nombre:
            self._set_estado(u"El nombre es obligatorio.")
            self._txt_nombre.Focus()
            return
        if not self._editing_item:
            return

        email = self._txt_email.Text.strip()
        self._editing_item.Nombre = nombre
        self._editing_item.Email  = email

        # Forzar refresco visual del DataGrid
        idx = self._personas.IndexOf(self._editing_item)
        self._personas.RemoveAt(idx)
        self._personas.Insert(idx, self._editing_item)

        _save_personas(self._personas, self._issues_dir, self._personas_file)
        self._editing_item = None
        self._txt_nombre.Text = ""
        self._txt_email.Text  = ""
        self._btn_agregar.Visibility    = System.Windows.Visibility.Visible
        self._btn_actualizar.Visibility = System.Windows.Visibility.Collapsed
        self._set_estado(u'"{}" actualizado correctamente.'.format(nombre))

    # ── Eliminar ─────────────────────────────────────────────────────────────
    def _on_eliminar_click(self, sender, args):
        persona = sender.Tag
        if not isinstance(persona, PersonaItem):
            return
        resp = MessageBox.Show(
            u'Eliminar a "{}" del directorio?'.format(persona.Nombre),
            u"Confirmar eliminacion",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question,
        )
        if str(resp) != "Yes":
            return
        if self._editing_item is persona:
            self._editing_item = None
            self._txt_nombre.Text = ""
            self._txt_email.Text  = ""
            self._btn_agregar.Visibility    = System.Windows.Visibility.Visible
            self._btn_actualizar.Visibility = System.Windows.Visibility.Collapsed
        self._personas.Remove(persona)
        _save_personas(self._personas, self._issues_dir, self._personas_file)
        self._set_estado(u'"{}" eliminado del directorio.'.format(persona.Nombre))

    def _on_key_down(self, sender, args):
        try:
            if args.Key == Key.Escape:
                self.Close()
        except Exception:
            pass

    # ── Estado ────────────────────────────────────────────────────────────────
    def _set_estado(self, msg):
        self._txt_estado.Text = msg


# ── Entry point ───────────────────────────────────────────────────────────────
try:
    w = PersonasWindow()
    w.ShowDialog()
except Exception as ex:
    try:
        from pyrevit import forms as pf
        pf.alert(u"Error al abrir Personas:\n\n{}".format(str(ex)), title="Error")
    except Exception:
        from System.Windows import MessageBox as MB
        MB.Show(str(ex), "Error — Personas")
